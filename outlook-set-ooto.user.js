// ==UserScript==
// @name         Outlook Set OOTO
// @namespace    https://amazon.com/
// @version      0.1
// @description  Set Out of Office directly on Outlook
// @author       @mofila
// @match        https://outlook.office.com/*
// @grant        GM.xmlHttpRequest
// @grant        unsafeWindow
// @downloadURL  https://raw.githubusercontent.com/Mofi-l/outlook-ooto-script/main/outlook-set-ooto.user.js
// @updateURL    https://raw.githubusercontent.com/Mofi-l/outlook-ooto-script/main/outlook-set-ooto.user.js
// ==/UserScript==

/* globals moment */

(function() {
    'use strict';

    // To mark script as installed
    localStorage.setItem('ooto_script_installed', 'true');
    
    // ============================================
    // OUTLOOK DIRECT PROVIDER
    // ============================================
    const OutlookDirectProvider = {
        name: 'Outlook Direct Provider',
        endpoint: 'https://outlook.office.com',

        // Get token from localStorage
        getToken() {
            // Look for the MSAL token with outlook.office.com scope
            for (let i = 0; i < localStorage.length; i++) {
                const key = localStorage.key(i);
                // Find the access token for outlook.office.com with calendar permissions
                if (key.includes('msal') &&
                    key.includes('accesstoken') &&
                    key.includes('outlook.office.com') &&
                    key.includes('calendars.readwrite')) {

                    const tokenData = localStorage.getItem(key);
                    if (tokenData) {
                        try {
                            const parsed = JSON.parse(tokenData);
                            // MSAL tokens are stored as JSON objects with a 'secret' property
                            if (parsed.secret) {
                                console.log('✅ Found Outlook API token');
                                return parsed.secret;
                            }
                        } catch (e) {
                            // If not JSON, might be the raw token
                            return tokenData;
                        }
                    }
                }
            }

            throw new Error('Could not find Outlook API token. Please refresh the page and try again.');
        },

        // Extract name from page DOM - simple approach
        getUserNameFromPage() {
            // Method 1: Try to get name from the account details section
            const nameElement = document.querySelector('#mectrl_currentAccount_primary, .mectrl_name');
            if (nameElement && nameElement.textContent.trim()) {
                return nameElement.textContent.trim();
            }

            // Method 2: Search for name pattern in all divs
            const allDivs = document.querySelectorAll('div');
            for (let div of allDivs) {
                const text = div.textContent.trim();
                // Look for name pattern (e.g., "Mofil, Abdul")
                if (/^[A-Z][a-z]+,\s[A-Z][a-z]+$/.test(text)) {
                    console.log('✅ Found name in div:', text);
                    return text;
                }
            }

            // Method 3: Extract from email as fallback
            try {
                const email = this.getUserEmailFromPage();
                const username = email.split('@');
                // Capitalize first letter
                const name = username.charAt(0).toUpperCase() + username.slice(1);
                console.log('✅ Extracted name from email:', name);
                return name;
            } catch (e) {
                return 'User'; // Final fallback
            }
        },

        // Extract email from page DOM
        getUserEmailFromPage() {
            // Method 1: Try the standard account details element
            const emailElement = document.querySelector('#mectrl_currentAccount_secondary');
            if (emailElement && emailElement.textContent.trim()) {
                return emailElement.textContent.trim();
            }

            // Method 2: Search for span with email pattern
            const allSpans = document.querySelectorAll('span');
            for (let span of allSpans) {
                const text = span.textContent.trim();
                // Look for Amazon email pattern
                if (/@amazon\.com$/.test(text) && text.length < 50) {
                    console.log('✅ Found email in span:', text);
                    return text;
                }
            }

            // Method 3: Search entire document for email pattern
            const bodyText = document.body.innerHTML;
            const emailMatch = bodyText.match(/([a-zA-Z0-9._-]+@amazon\.com)/);
            if (emailMatch) {
                console.log('✅ Found email in HTML:', emailMatch);
                return emailMatch;
            }

            throw new Error('Could not find user email on page. Please ensure you are logged in.');
        },

        async createOOTOMeetings({ organizer, organizerName, start, end, emailContent }) {
            const token = await this.getToken(); // Now returns a promise
            const subject = `OOTO | ${organizerName} | (${start.toLocaleDateString()} - ${end.toLocaleDateString()})`;

            const headers = {
                'authorization': `Bearer ${token}`,
                'content-type': 'application/json; charset=utf-8',
                'action': 'CreateItem',
                'x-owa-actionname': 'CreateCalendarItemAction'
            };

            return Promise.all([
                // Team meeting (Free) for hyd-microsites@amazon.com
                this.createMeeting(headers, {
                    __type: 'CalendarItem:#Exchange',
                    Subject: subject,
                    Body: { BodyType: 'HTML', Value: emailContent },
                    Sensitivity: 'Normal',
                    IsResponseRequested: false,
                    Start: start.toISOString(),
                    End: end.toISOString(),
                    FreeBusyType: 'Free',
                    RequiredAttendees: [{
                        __type: 'AttendeeType:#Exchange',
                        Mailbox: {
                            EmailAddress: 'hyd-microsites@amazon.com',
                            RoutingType: 'SMTP',
                            MailboxType: 'Mailbox',
                            OriginalDisplayName: 'hyd-microsites@amazon.com'
                        }
                    }]
                }),

                // Self meeting (OOF)
                this.createMeeting(headers, {
                    __type: 'CalendarItem:#Exchange',
                    Subject: subject,
                    Body: { BodyType: 'HTML', Value: emailContent },
                    Sensitivity: 'Normal',
                    IsResponseRequested: false,
                    Start: start.toISOString(),
                    End: end.toISOString(),
                    FreeBusyType: 'OOF',
                    RequiredAttendees: [{
                        __type: 'AttendeeType:#Exchange',
                        Mailbox: {
                            EmailAddress: organizer,
                            RoutingType: 'SMTP',
                            MailboxType: 'Mailbox',
                            OriginalDisplayName: organizer
                        }
                    }]
                })
            ]);
        },

        createMeeting(headers, meetingData) {
            return new Promise((resolve, reject) => {
                GM.xmlHttpRequest({
                    method: 'POST',
                    url: `${this.endpoint}/owa/service.svc`,
                    headers: headers,
                    data: JSON.stringify({
                        Header: {
                            RequestServerVersion: 'Exchange2013',
                            TimeZoneContext: { TimeZoneDefinition: { Id: 'UTC' } }
                        },
                        Body: {
                            Items: [meetingData],
                            SendMeetingInvitations: 'SendToAllAndSaveCopy'
                        }
                    }),
                    onerror: error => {
                        console.error('❌ Request error:', error);
                        reject(error);
                    },
                    onload: response => {
                        if (response.status === 200 || response.status === 201) {
                            console.log('✅ Meeting created successfully');
                            resolve();
                        } else {
                            console.error('❌ Error:', response.status, response.responseText);
                            reject(new Error(`Failed: ${response.status}`));
                        }
                    }
                });
            });
        }
    };

    // ============================================
    // EMAIL CONTENT DIALOG
    // ============================================
    function showEmailContentDialog(emailContent, callback) {
        const plainTextContent = emailContent
        .replace(/<br>/g, '')
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/g, ' ')
        .trim();

        const dialogHTML = `
        <div id="email-content-dialog" style="
            display: none;
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            padding: 20px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            z-index: 10001;
            width: 600px;">
            <h3 style="margin-top: 0; color: #232f3e;">Edit Email Content</h3>
            <textarea id="email-content-editor" style="
                width: 100%;
                height: 300px;
                margin: 10px 0;
                padding: 10px;
                border: 1px solid #ccc;
                border-radius: 4px;
                resize: vertical;
                font-family: Arial, sans-serif;
                font-size: 14px;
                line-height: 1.4;">${plainTextContent}</textarea>
            <div style="text-align: right; margin-top: 10px;">
                <button id="cancel-edit" style="
                    margin-right: 10px;
                    padding: 8px 16px;
                    background: #ff6b6b;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;">Cancel</button>
                <button id="send-without-edit" style="
                    margin-right: 10px;
                    padding: 8px 16px;
                    background: #f0f0f0;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;">Send Without Edit</button>
                <button id="save-and-send" style="
                    padding: 8px 16px;
                    background: #0066cc;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;">Save & Send</button>
            </div>
        </div>
        <div id="email-content-overlay" style="
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            z-index: 10000;">
        </div>`;

        document.body.insertAdjacentHTML('beforeend', dialogHTML);

        const dialog = document.getElementById('email-content-dialog');
        const overlay = document.getElementById('email-content-overlay');
        const editor = document.getElementById('email-content-editor');

        dialog.style.display = 'block';
        overlay.style.display = 'block';

        const closeDialog = () => {
            dialog.remove();
            overlay.remove();
        };

        document.getElementById('cancel-edit').addEventListener('click', () => {
            closeDialog();
            callback(null);
        });

        document.getElementById('send-without-edit').addEventListener('click', () => {
            closeDialog();
            callback(emailContent);
        });

        document.getElementById('save-and-send').addEventListener('click', () => {
            const editedContent = editor.value
            .split(String.fromCharCode(10))
            .map(line => line.trim())
            .join('<br>');
            closeDialog();
            callback(editedContent);
        });

    }

    function getEmailContent(userName) {
        return new Promise((resolve) => {
            const emailContent = `Hello Team,<br><br>
I will be out of office on the scheduled dates with no access to outlook, slack and chime.<br><br>
For project related queries, please reach out to my (enter supervisor login @).<br><br>
Regards,<br>
${userName}<br>
=====================================================`;

            showEmailContentDialog(emailContent, (content) => {
                resolve(content);
            });
        });
    }

    // ============================================
    // CREATE OOTO FORM (SIMPLIFIED - NO NAME/EMAIL FIELDS)
    // ============================================
    function createOOTOButton() {
        // Create floating button
        const buttonHTML = `
<div id="ooto-floating-button" style="
    position: fixed;
    bottom: 20px;
    right: 20px;
    z-index: 9999;">
    <button id="open-ooto-form" style="
        background: linear-gradient(135deg, #049796, #049796);
        color: white;
        padding: 12px 20px;
        border: none;
        border-radius: 25px;
        font-size: 14px;
        cursor: pointer;
        box-shadow: 0 4px 15px rgba(4, 151, 150, 0.4);
        transition: all 0.3s ease;
        font-weight: 500;">
        📅 Set Out of Office
    </button>
</div>

<div id="ooto-form" style="
    display: none;
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    padding: 30px;
    background: linear-gradient(135deg, #eaeded, #eaeded);
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5);
    z-index: 10000;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    width: 450px;">
    <h2 style="
        margin: 0 0 20px 0;
        font-size: 24px;
        color: #232f3e;
        text-align: center;
        font-weight: 600;">
        Set Out of Office
    </h2>

    <div style="margin-bottom: 15px;">
        <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block; font-weight: 500;">Start Date & Time:</label>
        <input type="datetime-local" id="ooto-start-datetime" style="
            width: 100%;
            padding: 12px;
            background: white;
            border: 1px solid #ccc;
            border-radius: 8px;
            font-size: 14px;
            outline: none;
            box-sizing: border-box;" />
    </div>

    <div style="margin-bottom: 20px;">
        <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block; font-weight: 500;">End Date & Time:</label>
        <input type="datetime-local" id="ooto-end-datetime" style="
            width: 100%;
            padding: 12px;
            background: white;
            border: 1px solid #ccc;
            border-radius: 8px;
            font-size: 14px;
            outline: none;
            box-sizing: border-box;" />
    </div>

    <div style="text-align: center;">
        <button id="submit-ooto" style="
            background: linear-gradient(135deg, #1db954, #1ed760);
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 10px;
            font-size: 14px;
            cursor: pointer;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);
            font-weight: 500;
            transition: all 0.3s ease;">
            Submit
        </button>
        <button id="cancel-ooto" style="
            background: linear-gradient(135deg, #ff6a6a, #ff3e3e);
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 10px;
            font-size: 14px;
            margin-left: 10px;
            cursor: pointer;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);
            font-weight: 500;
            transition: all 0.3s ease;">
            Cancel
        </button>
    </div>
</div>
<div id="ooto-overlay" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 9999;"></div>`;

        document.body.insertAdjacentHTML('beforeend', buttonHTML);

        const form = document.getElementById('ooto-form');
        const overlay = document.getElementById('ooto-overlay');
        const floatingButton = document.getElementById('open-ooto-form');

        // Hover effect for floating button
        floatingButton.addEventListener('mouseenter', () => {
            floatingButton.style.transform = 'scale(1.05)';
            floatingButton.style.boxShadow = '0 6px 20px rgba(4, 151, 150, 0.6)';
        });

        floatingButton.addEventListener('mouseleave', () => {
            floatingButton.style.transform = 'scale(1)';
            floatingButton.style.boxShadow = '0 4px 15px rgba(4, 151, 150, 0.4)';
        });

        // Open form
        floatingButton.addEventListener('click', () => {
            form.style.display = 'block';
            overlay.style.display = 'block';
        });

        // Cancel button
        document.getElementById('cancel-ooto').addEventListener('click', () => {
            form.style.display = 'none';
            overlay.style.display = 'none';
        });

        // Submit button - extracts user info from page DOM
        document.getElementById('submit-ooto').addEventListener('click', async () => {
            try {
                const startDatetime = document.getElementById('ooto-start-datetime').value;
                const endDatetime = document.getElementById('ooto-end-datetime').value;

                if (!startDatetime || !endDatetime) {
                    alert('Please fill in start and end date/time.');
                    return;
                }

                const submitBtn = document.getElementById('submit-ooto');
                const originalText = submitBtn.textContent;
                submitBtn.textContent = 'Processing...';
                await OutlookDirectProvider.getToken();
                submitBtn.disabled = true;

                // Extract user info from page DOM - simple and reliable
                const userName = OutlookDirectProvider.getUserNameFromPage();
                const userEmail = OutlookDirectProvider.getUserEmailFromPage();

                console.log('✅ Retrieved user info:', userName, userEmail);

                const start = new Date(startDatetime);
                const end = new Date(endDatetime);

                // Get email content
                const emailContent = await getEmailContent(userName);

                if (emailContent === null) {
                    submitBtn.textContent = originalText;
                    submitBtn.disabled = false;
                    return; // User cancelled
                }

                submitBtn.textContent = 'Creating meetings...';

                await OutlookDirectProvider.createOOTOMeetings({
                    organizer: userEmail,
                    organizerName: userName,
                    start,
                    end,
                    emailContent
                });

                alert('✅ Out of Office set successfully!');
                form.style.display = 'none';
                overlay.style.display = 'none';

                submitBtn.textContent = originalText;
                submitBtn.disabled = false;
            } catch (err) {
                console.error('❌ Error:', err);
                alert('Error: ' + err.message);

                const submitBtn = document.getElementById('submit-ooto');
                submitBtn.textContent = 'Submit';
                submitBtn.disabled = false;
            }
        });
    }

    // ============================================
    // INITIALIZE
    // ============================================
    // Wait for page to fully load
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            setTimeout(createOOTOButton, 1000); // Wait 1 second after DOM loads
        });
    } else {
        setTimeout(createOOTOButton, 1000); // Wait 1 second
    }

    console.log('✅ Outlook OOTO button initialized');

    // Debug: Check all localStorage keys for tokens
    console.log('🔍 Checking localStorage for tokens...');
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key.toLowerCase().includes('token') || key.toLowerCase().includes('auth')) {
            console.log(`Found: ${key}`);
        }
    }

})();
