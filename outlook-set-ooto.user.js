// ==UserScript==
// @name         Outlook Set OOTO
// @namespace    https://amazon.com/
// @version      0.3
// @description  Auto-create OOTO from Aura parameters
// @author       @mofila
// @match        https://outlook.office.com/*
// @grant        unsafeWindow
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM.xmlHttpRequest
// @grant        GM_xmlhttpRequest
// @downloadURL  https://raw.githubusercontent.com/Mofi-l/outlook-ooto-script/main/outlook-set-ooto.user.js
// @updateURL    https://raw.githubusercontent.com/Mofi-l/outlook-ooto-script/main/outlook-set-ooto.user.js
// ==/UserScript==

/* globals moment */

(function() {
    'use strict';

    // REQUIRED: Set installation marker for Aura to detect
    window.OUTLOOK_OOTO_SCRIPT_INSTALLED = true;
    console.log('[Outlook OOTO] Script loaded and marker set');

    // Version check configuration
    const CURRENT_VERSION = '0.3';
    const VERSION_CHECK_URL = 'https://raw.githubusercontent.com/Mofi-l/outlook-ooto-script/main/version.json';
    const SCRIPT_INSTALL_URL = 'https://raw.githubusercontent.com/Mofi-l/outlook-ooto-script/main/outlook-set-ooto.user.js';

    // Check for updates (non-blocking)
    function checkForUpdates() {
        const lastCheck = GM_getValue('last_version_check', 0);
        const now = Date.now();
        const oneDayMs = 24 * 60 * 60 * 1000;

        console.log('🔍 Checking for updates...');
        GM_setValue('last_version_check', now);

        const gmRequest = typeof GM !== 'undefined' && GM.xmlHttpRequest ? GM.xmlHttpRequest : GM_xmlhttpRequest;

        gmRequest({
            method: 'GET',
            url: VERSION_CHECK_URL,
            timeout: 5000, // Add timeout to prevent hanging
            onload: function(response) {
                if (response.status === 200) {
                    try {
                        const versionInfo = JSON.parse(response.responseText);
                        console.log('📦 Latest version:', versionInfo.version, 'Current:', CURRENT_VERSION);

                        // Compare versions properly (e.g., 0.3 > 0.2)
                        const latestVersion = parseFloat(versionInfo.version);
                        const currentVersion = parseFloat(CURRENT_VERSION);

                        if (latestVersion > currentVersion) {
                            console.log('🆕 New version available!');
                            showUpdateNotification(versionInfo);
                        } else {
                            console.log('✅ You have the latest version');
                        }
                    } catch (e) {
                        console.error('❌ Error parsing version info:', e);
                    }
                }
            },
            onerror: function(error) {
                console.log('❌ Update check failed:', error);
            },
            ontimeout: function() {
                console.log('⏱️ Update check timed out');
            }
        });
    }

    // Show update notification
    function showUpdateNotification(versionInfo) {
        // Remove any existing notification first
        const existingNotification = document.getElementById('ooto-update-notification');
        if (existingNotification) {
            existingNotification.remove();
        }

        const notificationHTML = `
    <div id="ooto-update-notification" style="
        position: fixed;
        top: 20px;
        right: 20px;
        background: linear-gradient(135deg, #ff6b6b, #ff8e53);
        color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
        z-index: 10002;
        max-width: 350px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <div style="display: flex; align-items: center; margin-bottom: 10px;">
            <span style="font-size: 24px; margin-right: 10px;">🔔</span>
            <strong style="font-size: 16px;">Update Available!</strong>
        </div>
        <p style="margin: 10px 0; font-size: 14px;">
            Version ${versionInfo.version} is now available.<br>
            <small>Current: ${CURRENT_VERSION}</small>
        </p>
        <p style="margin: 10px 0; font-size: 13px; opacity: 0.9;">
            ${versionInfo.description || 'New features and improvements'}
        </p>
        <div style="margin-top: 15px;">
            <button id="update-now-btn" style="
                background: white;
                color: #ff6b6b;
                padding: 10px 20px;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                font-weight: 600;
                margin-right: 10px;
                font-size: 14px;">
                Update Now
            </button>
            <button id="dismiss-update-btn" style="
                background: rgba(255, 255, 255, 0.2);
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                font-size: 14px;">
                Dismiss
            </button>
        </div>
    </div>`;

        document.body.insertAdjacentHTML('beforeend', notificationHTML);

        // Use event delegation
        document.addEventListener('click', function handleUpdateClick(e) {
            if (e.target.id === 'update-now-btn') {
                console.log('✅ Opening update URL:', SCRIPT_INSTALL_URL);
                window.open(SCRIPT_INSTALL_URL, '_blank');
                document.getElementById('ooto-update-notification')?.remove();
                document.removeEventListener('click', handleUpdateClick);
            } else if (e.target.id === 'dismiss-update-btn') {
                console.log('✅ Dismissing update notification');
                document.getElementById('ooto-update-notification')?.remove();
                document.removeEventListener('click', handleUpdateClick);
            }
        });
    }

    // Initialize - CRITICAL: Only call once
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            setTimeout(() => {
                checkForUpdates();
            }, 1000);
        });
    } else {
        setTimeout(() => {
            checkForUpdates();
        }, 1000);
    }

    // Set cross-domain marker that Aura can detect
    if (typeof GM_setValue !== 'undefined') {
        GM_setValue('ooto_script_installed', 'true');
    }

    // Also set localStorage as backup (for same-domain detection)
    localStorage.setItem('ooto_script_installed', 'true');

    GM_setValue('ooto_script_installed', 'true');

    // ============================================
    // OUTLOOK DIRECT PROVIDER
    // ============================================
    const OutlookDirectProvider = {
        name: 'Outlook Direct Provider',
        endpoint: 'https://outlook.office.com',
        isReady: false,
        userInfo: null,

        // Initialize provider
        async initialize() {
            try {
                console.log('🔧 Initializing OutlookDirectProvider...');
                this.userInfo = await this.getUserInfoFromOutlookAPI();
                this.isReady = true;
                console.log('✅ OutlookDirectProvider ready');
            } catch (error) {
                console.warn('⚠️ Could not get user info from API, will use page extraction:', error);
                // Fallback to page extraction
                try {
                    const email = this.getUserEmailFromPage();
                    const name = this.getUserNameFromPage();
                    this.userInfo = { email, displayName: name };
                    this.isReady = true;
                    console.log('✅ OutlookDirectProvider ready (using page data)');
                } catch (pageError) {
                    console.error('❌ Failed to initialize:', pageError);
                    this.isReady = false;
                }
            }
        },

        // Get user email (convenience method)
        getUserEmail() {
            if (!this.isReady || !this.userInfo) {
                // Fallback to page extraction if not initialized
                return this.getUserEmailFromPage();
            }
            return this.userInfo.email;
        },

        // Get user name (convenience method)
        getUserName() {
            if (!this.isReady || !this.userInfo) {
                // Fallback to page extraction if not initialized
                return this.getUserNameFromPage();
            }
            return this.userInfo.displayName;
        },

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

        async getUserInfoFromOutlookAPI() {
            const token = this.getToken();

            const gmRequest = typeof GM !== 'undefined' && GM.xmlHttpRequest ? GM.xmlHttpRequest : GM_xmlhttpRequest;

            return new Promise((resolve, reject) => {
                gmRequest({
                    method: 'GET',
                    url: 'https://outlook.office.com/api/v2.0/me',
                    headers: {
                        'Authorization': `Bearer ${token}`,
                        'Content-Type': 'application/json'
                    },
                    onload: response => {
                        if (response.status === 200) {
                            const userData = JSON.parse(response.responseText);
                            console.log('✅ Retrieved user info from Outlook API:', userData);
                            resolve({
                                displayName: userData.DisplayName,
                                email: userData.EmailAddress
                            });
                        } else {
                            console.error('❌ Error:', response.status);
                            reject(new Error(`Failed: ${response.status}`));
                        }
                    },
                    onerror: error => reject(error)
                });
            });
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
                console.log('✅ Found email in HTML:', emailMatch[1]);
                return emailMatch[1];
            }

            throw new Error('Could not find user email on page. Please ensure you are logged in.');
        },

        // Extract user name from page DOM
        getUserNameFromPage() {
            // Method 1: Try the standard account details element
            const nameElement = document.querySelector('#mectrl_currentAccount_primary');
            if (nameElement && nameElement.textContent.trim()) {
                return nameElement.textContent.trim();
            }

            // Method 2: Try to get from email (fallback)
            try {
                const email = this.getUserEmailFromPage();
                const username = email.split('@')[0];
                return username.charAt(0).toUpperCase() + username.slice(1);
            } catch (e) {
                return 'User';
            }
        },

        // Add this helper function to your OutlookDirectProvider object
        getLocalTimeZone() {
            // Get the user's timezone from their browser
            const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

            // Map IANA timezone names to Windows timezone IDs (required by Exchange)
            const timezoneMap = {
                // Americas
                'America/New_York': 'Eastern Standard Time',
                'America/Chicago': 'Central Standard Time',
                'America/Denver': 'Mountain Standard Time',
                'America/Los_Angeles': 'Pacific Standard Time',
                'America/Phoenix': 'US Mountain Standard Time',
                'America/Anchorage': 'Alaskan Standard Time',
                'America/Honolulu': 'Hawaiian Standard Time',

                // Europe
                'Europe/London': 'GMT Standard Time',
                'Europe/Paris': 'W. Europe Standard Time',
                'Europe/Berlin': 'W. Europe Standard Time',
                'Europe/Rome': 'W. Europe Standard Time',
                'Europe/Madrid': 'Romance Standard Time',
                'Europe/Amsterdam': 'W. Europe Standard Time',
                'Europe/Brussels': 'Romance Standard Time',
                'Europe/Vienna': 'W. Europe Standard Time',
                'Europe/Warsaw': 'Central European Standard Time',
                'Europe/Moscow': 'Russian Standard Time',

                // Asia
                'Asia/Kolkata': 'India Standard Time',
                'Asia/Shanghai': 'China Standard Time',
                'Asia/Hong_Kong': 'China Standard Time',
                'Asia/Tokyo': 'Tokyo Standard Time',
                'Asia/Seoul': 'Korea Standard Time',
                'Asia/Singapore': 'Singapore Standard Time',
                'Asia/Dubai': 'Arabian Standard Time',
                'Asia/Bangkok': 'SE Asia Standard Time',
                'Asia/Jakarta': 'SE Asia Standard Time',

                // Australia
                'Australia/Sydney': 'AUS Eastern Standard Time',
                'Australia/Melbourne': 'AUS Eastern Standard Time',
                'Australia/Brisbane': 'E. Australia Standard Time',
                'Australia/Perth': 'W. Australia Standard Time',

                // Others
                'Pacific/Auckland': 'New Zealand Standard Time',
                'Africa/Johannesburg': 'South Africa Standard Time',
                'America/Sao_Paulo': 'E. South America Standard Time'
            };

            // Return the Windows timezone ID, or default to UTC if not found
            const windowsTimeZone = timezoneMap[timeZone] || 'UTC';
            console.log(`✅ Detected timezone: ${timeZone} → ${windowsTimeZone}`);
            return windowsTimeZone;
        },

        async createOOTOMeetings({ organizer, organizerName, start, end, emailContent, subject }) {
            const token = await this.getToken();

            const headers = {
                'authorization': `Bearer ${token}`,
                'content-type': 'application/json; charset=utf-8',
                'action': 'CreateItem',
                'x-owa-actionname': 'CreateCalendarItemAction'
            };

            return Promise.all([
                // Meeting invite for all-microsites@amazon.com (Free)
                this.createMeetingWithInvite(headers, {
                    __type: 'CalendarItem:#Exchange',
                    Subject: subject,
                    Body: { BodyType: 'HTML', Value: emailContent },
                    Sensitivity: 'Normal',
                    IsResponseRequested: false,
                    Start: start.toISOString(),
                    End: end.toISOString(),
                    FreeBusyType: 'Free',
                    ReminderIsSet: false,
                    ReminderMinutesBeforeStart: 0,
                    RequiredAttendees: [{
                        __type: 'AttendeeType:#Exchange',
                        Mailbox: {
                            EmailAddress: 'all-microsites@amazon.com',
                            RoutingType: 'SMTP',
                            MailboxType: 'Mailbox',
                            OriginalDisplayName: 'all-microsites@amazon.com'
                        }
                    }]
                }),

                // Self calendar block (OOF) - no invite sent
                this.createMeetingNoInvite(headers, {
                    __type: 'CalendarItem:#Exchange',
                    Subject: subject,
                    Body: { BodyType: 'HTML', Value: emailContent },
                    Sensitivity: 'Normal',
                    IsResponseRequested: false,
                    Start: start.toISOString(),
                    End: end.toISOString(),
                    FreeBusyType: 'OOF'
                    // No RequiredAttendees - this creates a calendar block without inviting anyone
                })
            ]);
        },

        // Method for creating meeting WITH invite (for mofila@amazon.com)
        createMeetingWithInvite(headers, meetingData) {
            const gmRequest = typeof GM !== 'undefined' && GM.xmlHttpRequest ? GM.xmlHttpRequest : GM_xmlhttpRequest;

            return new Promise((resolve, reject) => {
                gmRequest({
                    method: 'POST',
                    url: `${this.endpoint}/owa/service.svc`,
                    headers: headers,
                    data: JSON.stringify({
                        Header: {
                            RequestServerVersion: 'Exchange2013',
                            TimeZoneContext: { TimeZoneDefinition: { Id: this.getLocalTimeZone() } }
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
                            console.log('✅ Meeting invite sent successfully');
                            resolve();
                        } else {
                            console.error('❌ Error:', response.status, response.responseText);
                            reject(new Error(`Failed: ${response.status}`));
                        }
                    }
                });
            });
        },

        // Method for creating calendar block WITHOUT invite (for OOF block)
        createMeetingNoInvite(headers, meetingData) {
            const gmRequest = typeof GM !== 'undefined' && GM.xmlHttpRequest ? GM.xmlHttpRequest : GM_xmlhttpRequest;

            return new Promise((resolve, reject) => {
                gmRequest({
                    method: 'POST',
                    url: `${this.endpoint}/owa/service.svc`,
                    headers: headers,
                    data: JSON.stringify({
                        Header: {
                            RequestServerVersion: 'Exchange2013',
                            TimeZoneContext: { TimeZoneDefinition: { Id: this.getLocalTimeZone() } }
                        },
                        Body: {
                            Items: [meetingData],
                            SendMeetingInvitations: 'SendToNone' // Key difference - don't send invite
                        }
                    }),
                    onerror: error => {
                        console.error('❌ Request error:', error);
                        reject(error);
                    },
                    onload: response => {
                        if (response.status === 200 || response.status === 201) {
                            console.log('✅ OOF calendar block created successfully');
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
    function showEmailContentDialog(subject, emailContent, callback) {
        const plainTextContent = emailContent
        .replace(/<br>/g, '\n')
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
            <h3 style="margin-top: 0; color: #232f3e;">Edit Meeting Details</h3>

            <!-- Subject Line Field -->
            <div style="margin-bottom: 15px;">
                <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block; font-weight: 500;">Subject Line:</label>
                <input type="text" id="email-subject-editor" value="${subject}" style="
                    width: 100%;
                    padding: 10px;
                    border: 1px solid #ccc;
                    border-radius: 4px;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                    box-sizing: border-box;" />
            </div>

            <!-- Email Body Field -->
            <div style="margin-bottom: 15px;">
                <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block; font-weight: 500;">Email Body:</label>
                <textarea id="email-content-editor" style="
                    width: 100%;
                    height: 250px;
                    padding: 10px;
                    border: 1px solid #ccc;
                    border-radius: 4px;
                    resize: vertical;
                    font-family: Arial, sans-serif;
                    font-size: 14px;
                    line-height: 1.4;
                    box-sizing: border-box;">${plainTextContent}</textarea>
            </div>

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
        const subjectEditor = document.getElementById('email-subject-editor');

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
            callback({ subject, emailContent });
        });

        document.getElementById('save-and-send').addEventListener('click', () => {
            const editedSubject = subjectEditor.value.trim();
            const editedContent = editor.value
            .split('\n')
            .map(line => line.trim())
            .join('<br>');
            closeDialog();
            callback({ subject: editedSubject, emailContent: editedContent });
        });
    }


    function getEmailContent(userName, organizerName, start, end) {
        return new Promise((resolve) => {
            const subject = `OOTO | ${organizerName} | (${start.toLocaleDateString()} - ${end.toLocaleDateString()})`;

            const emailContent = `Hello Team,<br>
I will be out of office on the scheduled dates with no access to outlook, slack and chime.<br>
For project related queries, please reach out to my (@enter supervisor login).<br>
Regards,
${userName}
=====================================================`;

            showEmailContentDialog(subject, emailContent, (result) => {
                resolve(result);
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
        <label style="display: flex; align-items: center; font-size: 14px; color: #232f3e; cursor: pointer;">
            <input type="checkbox" id="ooto-all-day" style="
                margin-right: 8px;
                width: 18px;
                height: 18px;
                cursor: pointer;" />
            <span style="font-weight: 500;">All Day Event</span>
        </label>
    </div>

    <div style="margin-bottom: 15px;">
        <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block; font-weight: 500;">Start Date:</label>
        <input type="date" id="ooto-start-date" style="
            width: 100%;
            padding: 12px;
            background: white;
            border: 1px solid #ccc;
            border-radius: 8px;
            font-size: 14px;
            outline: none;
            box-sizing: border-box;" />
    </div>

    <div id="start-time-container" style="margin-bottom: 15px;">
        <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block; font-weight: 500;">Start Time:</label>
        <input type="time" id="ooto-start-time" style="
            width: 100%;
            padding: 12px;
            background: white;
            border: 1px solid #ccc;
            border-radius: 8px;
            font-size: 14px;
            outline: none;
            box-sizing: border-box;" />
    </div>

    <div style="margin-bottom: 15px;">
        <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block; font-weight: 500;">End Date:</label>
        <input type="date" id="ooto-end-date" style="
            width: 100%;
            padding: 12px;
            background: white;
            border: 1px solid #ccc;
            border-radius: 8px;
            font-size: 14px;
            outline: none;
            box-sizing: border-box;" />
    </div>

    <div id="end-time-container" style="margin-bottom: 20px;">
        <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block; font-weight: 500;">End Time:</label>
        <input type="time" id="ooto-end-time" style="
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
        const allDayCheckbox = document.getElementById('ooto-all-day');
        const startTimeContainer = document.getElementById('start-time-container');
        const endTimeContainer = document.getElementById('end-time-container');

        // Handle All Day checkbox toggle
        allDayCheckbox.addEventListener('change', () => {
            if (allDayCheckbox.checked) {
                // Hide time inputs when All Day is checked
                startTimeContainer.style.display = 'none';
                endTimeContainer.style.display = 'none';
            } else {
                // Show time inputs when All Day is unchecked
                startTimeContainer.style.display = 'block';
                endTimeContainer.style.display = 'block';
            }
        });

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

        // Submit button - updated to handle All Day events
        document.getElementById('submit-ooto').addEventListener('click', async () => {
            try {
                const isAllDay = allDayCheckbox.checked;
                const startDate = document.getElementById('ooto-start-date').value;
                const endDate = document.getElementById('ooto-end-date').value;

                if (!startDate || !endDate) {
                    alert('Please fill in start and end dates.');
                    return;
                }

                let start, end;

                if (isAllDay) {
                    // For all-day events, set time to start of day (00:00) and end of day (23:59)
                    start = new Date(startDate + 'T00:00:00');
                    end = new Date(endDate + 'T23:59:59');
                } else {
                    const startTime = document.getElementById('ooto-start-time').value;
                    const endTime = document.getElementById('ooto-end-time').value;

                    if (!startTime || !endTime) {
                        alert('Please fill in start and end times.');
                        return;
                    }

                    start = new Date(startDate + 'T' + startTime);
                    end = new Date(endDate + 'T' + endTime);
                }

                const submitBtn = document.getElementById('submit-ooto');
                const originalText = submitBtn.textContent;
                submitBtn.textContent = 'Processing...';
                await OutlookDirectProvider.getToken();
                submitBtn.disabled = true;

                // AWAIT the API call to get user info
                const userInfo = await OutlookDirectProvider.getUserInfoFromOutlookAPI();
                const userName = userInfo.displayName;
                const userEmail = userInfo.email;

                console.log('✅ Retrieved user info:', userName, userEmail);

                // Get email content
                // Get email content with subject
                const result = await getEmailContent(userName, userName, start, end);

                if (result === null) {
                    submitBtn.textContent = originalText;
                    submitBtn.disabled = false;
                    return; // User cancelled
                }

                const { subject, emailContent } = result;

                submitBtn.textContent = 'Creating meetings...';

                await OutlookDirectProvider.createOOTOMeetings({
                    organizer: userEmail,
                    organizerName: userName,
                    start,
                    end,
                    emailContent,
                    subject,
                    isAllDay
                });

                alert('✅ Out of Office set successfully!');

                // Sync with Aura S3
                try {
                    await saveOOTOToS3({
                        organizer: userEmail,
                        startDate: start.toISOString().split('T')[0],
                        endDate: end.toISOString().split('T')[0],
                        startTime: isAllDay ? '00:00' : start.toTimeString().split(' ')[0].substring(0, 5),
                        endTime: isAllDay ? '23:59' : end.toTimeString().split(' ')[0].substring(0, 5)
                    });
                } catch (syncError) {
                    console.error('Aura sync error:', syncError);
                    // Don't fail the whole operation if sync fails
                }

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
    // AURA S3 SYNC FUNCTIONS
    // ============================================

    // Helper: Generate date range
    function getDateRange(startDate, endDate) {
        const dates = [];
        const current = new Date(startDate);
        const end = new Date(endDate);

        while (current <= end) {
            dates.push(current.toISOString().split('T')[0]);
            current.setDate(current.getDate() + 1);
        }

        return dates;
    }

    // Helper: Show notification
    function showNotification(message, type = 'success') {
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px 20px;
            background: ${type === 'success' ? '#10b981' : type === 'warning' ? '#f59e0b' : '#ef4444'};
            color: white;
            border-radius: 8px;
            z-index: 100000;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            font-size: 14px;
            max-width: 300px;
        `;
        notification.textContent = message;
        document.body.appendChild(notification);

        setTimeout(() => notification.remove(), 5000);
    }

    // ============================================
    // SDK LOADING FUNCTIONS (Using GM_xmlhttpRequest to bypass CSP)
    // ============================================

    function loadCognitoSdk() {
        return new Promise((resolve, reject) => {
            // Check if already loaded
            if (typeof AmazonCognitoIdentity !== 'undefined') {
                console.log('✅ Amazon Cognito SDK already loaded');
                resolve();
                return;
            }

            // Check unsafeWindow
            if (typeof unsafeWindow !== 'undefined' && unsafeWindow.AmazonCognitoIdentity) {
                window.AmazonCognitoIdentity = unsafeWindow.AmazonCognitoIdentity;
                console.log('✅ Found Cognito SDK in unsafeWindow');
                resolve();
                return;
            }

            console.log('📦 Loading Amazon Cognito SDK via GM_xmlhttpRequest...');

            const gmRequest = typeof GM_xmlhttpRequest !== 'undefined' ? GM_xmlhttpRequest :
                             (typeof GM !== 'undefined' && GM.xmlHttpRequest ? GM.xmlHttpRequest : null);

            if (!gmRequest) {
                console.error('❌ GM_xmlhttpRequest not available');
                reject(new Error('GM_xmlhttpRequest not available'));
                return;
            }

            gmRequest({
                method: 'GET',
                url: 'https://cdnjs.cloudflare.com/ajax/libs/amazon-cognito-identity-js/5.2.1/amazon-cognito-identity.min.js',
                onload: function(response) {
                    try {
                        // Inject the script into the page
                        const script = document.createElement('script');
                        script.textContent = response.responseText;
                        document.head.appendChild(script);

                        setTimeout(() => {
                            if (typeof AmazonCognitoIdentity !== 'undefined') {
                                console.log('✅ Amazon Cognito SDK loaded successfully');
                                resolve();
                            } else if (typeof unsafeWindow !== 'undefined' && unsafeWindow.AmazonCognitoIdentity) {
                                window.AmazonCognitoIdentity = unsafeWindow.AmazonCognitoIdentity;
                                console.log('✅ Found Cognito SDK in unsafeWindow after injection');
                                resolve();
                            } else {
                                console.error('❌ Cognito SDK injected but not accessible');
                                reject(new Error('Cognito SDK injected but not accessible'));
                            }
                        }, 500);
                    } catch (error) {
                        console.error('❌ Error injecting Cognito SDK:', error);
                        reject(error);
                    }
                },
                onerror: function(error) {
                    console.error('❌ Failed to fetch Cognito SDK:', error);
                    reject(new Error('Failed to fetch Cognito SDK'));
                }
            });
        });
    }

    async function loadAwsSdk() {
        return new Promise((resolve, reject) => {
            // Check if already loaded
            if (typeof AWS !== 'undefined' && AWS.S3) {
                console.log('✅ AWS SDK already loaded and AWS.S3 is available');
                resolve();
                return;
            }

            // Check unsafeWindow
            if (typeof unsafeWindow !== 'undefined' && unsafeWindow.AWS && unsafeWindow.AWS.S3) {
                window.AWS = unsafeWindow.AWS;
                console.log('✅ Found AWS SDK in unsafeWindow');
                resolve();
                return;
            }

            console.log('📦 Loading AWS SDK via GM_xmlhttpRequest...');

            const gmRequest = typeof GM_xmlhttpRequest !== 'undefined' ? GM_xmlhttpRequest :
                             (typeof GM !== 'undefined' && GM.xmlHttpRequest ? GM.xmlHttpRequest : null);

            if (!gmRequest) {
                console.error('❌ GM_xmlhttpRequest not available');
                reject(new Error('GM_xmlhttpRequest not available'));
                return;
            }

            gmRequest({
                method: 'GET',
                url: 'https://sdk.amazonaws.com/js/aws-sdk-2.1109.0.min.js',
                onload: function(response) {
                    try {
                        // Inject the script into the page
                        const script = document.createElement('script');
                        script.textContent = response.responseText;
                        document.head.appendChild(script);

                        setTimeout(() => {
                            if (typeof AWS !== 'undefined' && AWS.S3) {
                                console.log('✅ AWS SDK loaded successfully');
                                resolve();
                            } else if (typeof unsafeWindow !== 'undefined' && unsafeWindow.AWS && unsafeWindow.AWS.S3) {
                                window.AWS = unsafeWindow.AWS;
                                console.log('✅ Found AWS SDK in unsafeWindow after injection');
                                resolve();
                            } else {
                                console.error('❌ AWS SDK injected but not accessible');
                                reject(new Error('AWS SDK injected but not accessible'));
                            }
                        }, 500);
                    } catch (error) {
                        console.error('❌ Error injecting AWS SDK:', error);
                        reject(error);
                    }
                },
                onerror: function(error) {
                    console.error('❌ Failed to fetch AWS SDK:', error);
                    reject(new Error('Failed to fetch AWS SDK'));
                }
            });
        });
    }

    // ============================================
    // AUTHENTICATION
    // ============================================

    // Authenticate with Cognito (same as Aura)
    async function authenticateWithCognito() {
        return new Promise(async (resolve, reject) => {
            try {
                // Load Cognito SDK first
                await loadCognitoSdk();

                if (typeof AmazonCognitoIdentity === 'undefined') {
                    throw new Error('Cognito SDK failed to load');
                }

                // Try to get stored credentials first
                const storedUsername = GM_getValue('aura_username', '');
                const storedToken = GM_getValue('aura_token', '');
                const tokenExpiry = GM_getValue('aura_token_expiry', 0);

                // Check if token is still valid (within 1 hour)
                if (storedToken && Date.now() < tokenExpiry) {
                    console.log('✅ Using stored Aura credentials');
                    resolve(storedToken);
                    return;
                }

                // Get username from Outlook email automatically
                const userEmail = await OutlookDirectProvider.getUserEmail();
                const username = userEmail.split('@')[0];

                console.log(`🔐 Authenticating as: ${username}`);

                // Need password
                const password = prompt(`Enter Aura password for ${username}:`);
                if (!password) {
                    reject(new Error('Password required for Aura sync'));
                    return;
                }

                const authenticationData = {
                    Username: username,
                    Password: password,
                };

                const authenticationDetails = new AmazonCognitoIdentity.AuthenticationDetails(authenticationData);

                const poolData = {
                    UserPoolId: 'eu-north-1_V9kLPNVXl',
                    ClientId: '68caeoofa7hl7p7pvs65bb2hrv'
                };

                const userPool = new AmazonCognitoIdentity.CognitoUserPool(poolData);
                const userData = { Username: username, Pool: userPool };
                const cognitoUser = new AmazonCognitoIdentity.CognitoUser(userData);

                cognitoUser.authenticateUser(authenticationDetails, {
                    onSuccess: (result) => {
                        const token = result.getIdToken().getJwtToken();

                        // Store credentials for 1 hour
                        GM_setValue('aura_username', username);
                        GM_setValue('aura_token', token);
                        GM_setValue('aura_token_expiry', Date.now() + 3600000); // 1 hour

                        console.log('✅ Aura authentication successful');
                        resolve(token);
                    },
                    onFailure: (err) => {
                        console.error('❌ Aura authentication failed:', err);
                        reject(err);
                    }
                });
            } catch (error) {
                reject(error);
            }
        });
    }

    // Save OOTO to S3 for Aura sync
    async function saveOOTOToS3(meetingData) {
        try {
            console.log('💾 Saving OOTO to Aura S3...');
            showNotification('Syncing with Aura...', 'success');

            // Get username from email
            const username = meetingData.organizer.split('@')[0];

            // Load AWS SDK first
            await loadAwsSdk();

            // Authenticate with Cognito (this will load Cognito SDK)
            const token = await authenticateWithCognito();

            // Configure AWS
            AWS.config.update({
                region: 'eu-north-1',
                credentials: new AmazonCognitoIdentity.CognitoIdentityCredentials({
                    IdentityPoolId: 'eu-north-1:98c07095-e731-4219-bebe-db4dab892ea8',
                    Logins: {
                        'cognito-idp.eu-north-1.amazonaws.com/eu-north-1_V9kLPNVXl': token
                    }
                })
            });

            // Wait for credentials
            await new Promise((resolve, reject) => {
                AWS.config.credentials.get(err => {
                    if (err) reject(err);
                    else resolve();
                });
            });

            const s3 = new AWS.S3();

            // Generate date range
            const dates = getDateRange(meetingData.startDate, meetingData.endDate);

            console.log(`📅 Saving OOTO for ${dates.length} days...`);

            // Save OOTO for each date
            const savePromises = dates.map(date => {
                const ootoRecord = {
                    username: username,
                    date: date,
                    startDate: meetingData.startDate,
                    endDate: meetingData.endDate,
                    startTime: meetingData.startTime || '00:00',
                    endTime: meetingData.endTime || '23:59',
                    reason: 'Out of Office',
                    leaveType: 'planned',
                    source: 'outlook-ooto-script',
                    status: 'active',
                    createdAt: new Date().toISOString(),
                    createdBy: username
                };

                return s3.putObject({
                    Bucket: 'real-time-databucket',
                    Key: `ooto-status/${username}/${date}.json`,
                    Body: JSON.stringify(ootoRecord),
                    ContentType: 'application/json'
                }).promise();
            });

            await Promise.all(savePromises);

            console.log(`✅ OOTO saved to Aura S3 for ${dates.length} days`);
            showNotification(`✅ Synced with Aura! (${dates.length} days)`, 'success');

        } catch (error) {
            console.error('❌ Error saving OOTO to S3:', error);
            showNotification('⚠️ OOTO created in Outlook but failed to sync with Aura. You can mark it manually in Aura.', 'warning');
        }
    }

    // ============================================
    // INITIALIZE
    // ============================================

    // Create OOTO button and check for Aura parameters
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            console.log('📄 DOM loaded, initializing...');
            setTimeout(async () => {
                await OutlookDirectProvider.initialize();
                // Only check for Aura parameters, don't create button
                checkForAuraOOTOParameters();
            }, 1000);
        });
    } else {
        console.log('📄 DOM already loaded, initializing...');
        setTimeout(async () => {
            await OutlookDirectProvider.initialize();
            // Only check for Aura parameters, don't create button
            checkForAuraOOTOParameters();
        }, 1000);
    }

    console.log('✅ Outlook OOTO script loaded (v' + CURRENT_VERSION + ')');

    // ============================================
    // AUTO-CREATE OOTO FROM AURA PARAMETERS
    // ============================================

    // Check if URL has Aura OOTO parameters
    function checkForAuraOOTOParameters() {
        const urlParams = new URLSearchParams(window.location.search);

        if (urlParams.get('aura_ooto') === 'true') {
            console.log('🎯 Aura OOTO parameters detected!');

            const ootoParams = {
                startDate: urlParams.get('start_date'),
                endDate: urlParams.get('end_date'),
                isAllDay: urlParams.get('all_day') === '1',
                startTime: urlParams.get('start_time'),
                endTime: urlParams.get('end_time'),
                subject: urlParams.get('subject'),
                emailBody: urlParams.get('email_body'),
                username: urlParams.get('username')
            };

            console.log('📋 OOTO Parameters:', ootoParams);

            // Wait for page to be ready, then create OOTO
            setTimeout(() => {
                createOOTOFromAura(ootoParams);
            }, 5000); // Wait 5 seconds for Outlook to fully load
        }
    }

    // Create OOTO meetings from Aura parameters
    let ootoCreationInProgress = false; // Prevent duplicate calls

    async function createOOTOFromAura(params) {
        // Prevent duplicate execution
        if (ootoCreationInProgress) {
            console.log('⚠️ OOTO creation already in progress, skipping duplicate call');
            return;
        }

        ootoCreationInProgress = true;

        try {
            console.log('🚀 Creating OOTO from Aura parameters...');

            // Show notification
            showNotification('Creating OOTO meetings...', 'success');

            // Try to initialize if not ready
            if (!OutlookDirectProvider.isReady) {
                console.log('🔧 OutlookDirectProvider not ready, attempting to initialize...');
                try {
                    await OutlookDirectProvider.initialize();
                } catch (initError) {
                    console.warn('⚠️ Initialize failed, will retry:', initError);
                }
            }

            // Wait for OutlookDirectProvider to be ready
            let attempts = 0;
            while (!OutlookDirectProvider.isReady && attempts < 50) {
                console.log('⏳ Waiting for OutlookDirectProvider...', attempts);
                await new Promise(r => setTimeout(r, 200));
                attempts++;
            }

            if (!OutlookDirectProvider.isReady) {
                throw new Error('OutlookDirectProvider not ready after 50 attempts. Please refresh the page and try again.');
            }

            // Get user info
            const userEmail = await OutlookDirectProvider.getUserEmail();
            const userName = await OutlookDirectProvider.getUserName();

            console.log(`👤 User: ${userName} (${userEmail})`);

            // Parse dates
            let start, end;

            if (params.isAllDay) {
                start = new Date(params.startDate + 'T00:00:00');
                end = new Date(params.endDate + 'T23:59:59');
            } else {
                start = new Date(params.startDate + 'T' + params.startTime + ':00');
                end = new Date(params.endDate + 'T' + params.endTime + ':00');
            }

            console.log(`� Creating OOTO from ${start} to ${end}`);

            // Use email content from Aura parameters (no dialog needed)
            const subject = params.subject || `OOTO | ${userName} | (${start.toLocaleDateString()} - ${end.toLocaleDateString()})`;
            const emailContent = params.emailBody ? params.emailBody.replace(/\n/g, '<br>') : `Hello Team,<br>I will be out of office on the scheduled dates with no access to outlook, slack and chime.<br>For project related queries, please reach out to my supervisor.<br>Regards,<br>${userName}`;

            console.log('📧 Using email content from Aura:', { subject, emailContent });

            // Create OOTO meetings directly without showing dialog
            await OutlookDirectProvider.createOOTOMeetings({
                organizer: userEmail,
                organizerName: userName,
                start: start,
                end: end,
                emailContent: emailContent,
                subject: subject
            });

            console.log('✅ OOTO meetings created successfully!');
            showNotification('✅ OOTO created successfully!', 'success');

            // Send success message back to Aura via localStorage
            const successData = {
                success: true,
                startDate: params.startDate,
                endDate: params.endDate,
                timestamp: new Date().toISOString()
            };

            localStorage.setItem('aura_ooto_success', JSON.stringify(successData));
            console.log('📤 Success message sent to Aura');

            // Show success message to user
            setTimeout(() => {
                alert('✅ Out of Office created successfully!\n\nYou can close this tab and return to Aura.');
            }, 1000);

        } catch (error) {
            console.error('❌ Error creating OOTO from Aura:', error);
            showNotification('❌ Failed to create OOTO: ' + error.message, 'error');

            // Send error message back to Aura
            const errorData = {
                success: false,
                error: error.message,
                timestamp: new Date().toISOString()
            };

            localStorage.setItem('aura_ooto_error', JSON.stringify(errorData));

            alert('❌ Failed to create OOTO: ' + error.message + '\n\nPlease try again or create manually.');
        } finally {
            ootoCreationInProgress = false;
        }
    }

    // Check for Aura parameters on page load (only once)
    // Removed duplicate call - already called in DOMContentLoaded handler above

    // Debug: Check all localStorage keys for tokens
    console.log('🔍 Checking localStorage for tokens...');
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key.toLowerCase().includes('token') || key.toLowerCase().includes('auth')) {
            console.log(`Found: ${key}`);
        }
    }
})();
