/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { getUserData, getGraphToken, getMiddletierToken } from "../helpers/sso-helper";
import { callGetRandomMobile } from "../helpers/middle-tier-calls";
import { APP_VERSION, GIT_COMMIT } from "../helpers/version";

interface Recipient {
  no: number;
  email: string;
  mobile: string | null;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getProfileButton").onclick = run;
    document.getElementById("loadRecipientsButton").onclick = loadRecipients;
    document.getElementById("secureToggle").onclick = toggleSecureSend;
    
    // Initialize secure toggle button icon
    initializeSecureToggle();
    
    // Display version info
    displayVersionInfo();
    
    // Auto-load recipients when task pane opens (if composing)
    if (Office.context.mailbox.item) {
 //     loadRecipients();
    }
  }
});

export async function run() {
  getUserData(writeDataToOfficeDocument);
}

export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];
  let userProfileInfo: string[] = [];
  userProfileInfo.push(result["displayName"]);
  userProfileInfo.push(result["jobTitle"]);
  userProfileInfo.push(result["mail"]);
  userProfileInfo.push(result["mobilePhone"]);
  userProfileInfo.push(result["officeLocation"]);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}

// ============================================
// Recipients Functionality
// ============================================

async function loadRecipients() {
  showLoading(true);
  hideEmptyState();
  hideRecipientsList();

  try {
    const item = Office.context.mailbox.item;
    const allRecipients: { email: string; type: string }[] = [];

    // Get To recipients
    const toRecipients = await getRecipientsAsync(item.to);
    allRecipients.push(...toRecipients.map(email => ({ email, type: 'To' })));

    // Get Cc recipients
    const ccRecipients = await getRecipientsAsync(item.cc);
    allRecipients.push(...ccRecipients.map(email => ({ email, type: 'Cc' })));

    // Get Bcc recipients (if accessible)
    if (item.bcc) {
      const bccRecipients = await getRecipientsAsync(item.bcc);
      allRecipients.push(...bccRecipients.map(email => ({ email, type: 'Bcc' })));
    }

    if (allRecipients.length === 0) {
      showLoading(false);
      showEmptyState();
      return;
    }

    // Fetch mobile numbers from Microsoft Graph
    await fetchMobileNumbers(allRecipients);

  } catch (error) {
    console.error("Error loading recipients:", error);
    showLoading(false);
    showEmptyState();
  }
}

function getRecipientsAsync(recipientField: any): Promise<string[]> {
  return new Promise((resolve, reject) => {
    recipientField.getAsync((result: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emails = result.value.map(r => r.emailAddress);
        resolve(emails);
      } else {
        reject(result.error);
      }
    });
  });
}

async function fetchMobileNumbers(recipients: { email: string; type: string }[]) {
  try {
    // Get access token for Microsoft Graph
    const token = await getGraphToken();

    // First: fetch mobile numbers from Graph for each recipient
    let recipientsWithMobile = await Promise.all(
      recipients.map(async (recipient, index) => {
        const mobile = await getMobileNumber(recipient.email, token);
        return {
          no: index + 1,
          email: recipient.email,
          mobile: mobile
        };
      })
    );

    // Filter those without a mobile and call our middle-tier API to get a random mobile
    const missing = recipientsWithMobile.filter(r => !r.mobile);
    if (missing.length > 0) {
      try {
        // Get SSO token to present to middle-tier
        const middletierToken = await getMiddletierToken();

        const randomResults = await Promise.all(
          missing.map(async (r) => {
            try {
              const res = await callGetRandomMobile(middletierToken, r.email);
              return { email: r.email, mobile: res && res.mobile ? res.mobile : null };
            } catch (err) {
              // If the middle-tier call fails for a particular recipient, leave mobile null
              console.error(`Error fetching random mobile for ${r.email}:`, err);
              return { email: r.email, mobile: null };
            }
          })
        );

        // Merge random results back into recipientsWithMobile
        const randomMap = new Map(randomResults.map(rr => [rr.email, rr.mobile]));
        recipientsWithMobile = recipientsWithMobile.map(r => ({ ...r, mobile: r.mobile || randomMap.get(r.email) || null }));
      } catch (err) {
        console.error("Error getting middletier token or random mobiles:", err);
        // proceed with what we have (Graph results)
      }
    }

    displayRecipients(recipientsWithMobile);
  } catch (error) {
    console.error("Error fetching mobile numbers:", error);
    
    // Display recipients without mobile numbers
    const recipientsWithoutMobile = recipients.map((r, i) => ({
      no: i + 1,
      email: r.email,
      mobile: null
    }));
    
    displayRecipients(recipientsWithoutMobile);
  }
}

async function getMobileNumber(email: string, token: string): Promise<string | null> {
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}?$select=mobilePhone`,
      {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      }
    );

    if (response.ok) {
      const data = await response.json();
      return data.mobilePhone || null;
    }
  } catch (error) {
    console.error(`Error fetching mobile for ${email}:`, error);
  }
  return null;
}

function displayRecipients(recipients: Recipient[]) {
  showLoading(false);

  if (recipients.length === 0) {
    showEmptyState();
    return;
  }

  const tbody = document.getElementById("recipients-body");
  tbody.innerHTML = ""; // Clear existing rows

  recipients.forEach(recipient => {
    const row = document.createElement("div");
    row.className = "ig-list-row";
    
    // Create mobile input with validation
    const mobileInputId = `mobile-${recipient.no}`;
    const mobileValue = recipient.mobile || '';
    const isValid = mobileValue ? validateIsraeliMobile(mobileValue) : null;
    const statusIcon = getStatusIcon(isValid);
    
    row.innerHTML = `
      <div class="ig-list-cell ig-cell-no">${recipient.no}</div>
      <div class="ig-list-cell ig-cell-email" title="${recipient.email}">${recipient.email}</div>
      <div class="ig-list-cell ig-cell-mobile">
        <div class="ig-mobile-input-wrapper">
          <input 
            type="tel" 
            id="${mobileInputId}"
            class="ig-mobile-input ${isValid === false ? 'invalid' : ''}" 
            value="${mobileValue}"
            placeholder="05X-XXXXXXX"
            data-email="${recipient.email}"
          />
        </div>
      </div>
      <div class="ig-list-cell ig-cell-status">
        <i class="ms-Icon ms-Icon--${statusIcon.icon} ig-status-icon ${statusIcon.class}"></i>
      </div>
    `;
    
    tbody.appendChild(row);
    
    // Add event listener for validation
    const input = document.getElementById(mobileInputId) as HTMLInputElement;
    input.addEventListener('input', (e) => validateMobileInput(e.target as HTMLInputElement));
    input.addEventListener('blur', (e) => validateMobileInput(e.target as HTMLInputElement));
  });

  showRecipientsList();
}

// ============================================
// Israeli Mobile Number Validation
// ============================================

function validateIsraeliMobile(mobile: string): boolean {
  if (!mobile) return false;
  
  // Remove all non-digit characters
  const cleaned = mobile.replace(/\D/g, '');
  
  // Israeli mobile formats:
  // 05X-XXXXXXX (10 digits starting with 05)
  // 972-5X-XXXXXXX (12 digits starting with 972-5)
  // +972-5X-XXXXXXX (12 digits starting with +972-5)
  
  // Check for Israeli mobile: starts with 05 and has 10 digits
  if (cleaned.length === 10 && cleaned.startsWith('05')) {
    return true;
  }
  
  // Check for international format: +972-5X or 972-5X
  if (cleaned.length === 12 && cleaned.startsWith('9725')) {
    return true;
  }
  
  return false;
}

function validateMobileInput(input: HTMLInputElement) {
  const value = input.value.trim();
  const statusCell = input.closest('.ig-list-row').querySelector('.ig-cell-status');
  const statusIcon = statusCell.querySelector('.ig-status-icon');
  
  if (!value) {
    // Empty - neutral state
    input.classList.remove('invalid');
    statusIcon.className = 'ms-Icon ms-Icon--StatusCircleQuestionMark ig-status-icon empty';
    return;
  }
  
  const isValid = validateIsraeliMobile(value);
  
  if (isValid) {
    input.classList.remove('invalid');
    statusIcon.className = 'ms-Icon ms-Icon--CheckMark ig-status-icon valid';
  } else {
    input.classList.add('invalid');
    statusIcon.className = 'ms-Icon ms-Icon--StatusErrorFull ig-status-icon invalid';
  }
}

function getStatusIcon(isValid: boolean | null): { icon: string; class: string } {
  if (isValid === null) {
    return { icon: 'StatusCircleQuestionMark', class: 'empty' };
  }
  return isValid 
    ? { icon: 'CheckMark', class: 'valid' }
    : { icon: 'StatusErrorFull', class: 'invalid' };
}

// UI Helper Functions
function showLoading(show: boolean) {
  const loadingEl = document.getElementById("recipients-loading");
  loadingEl.style.display = show ? "flex" : "none";
}

function showEmptyState() {
  document.getElementById("no-recipients").style.display = "flex";
}

function hideEmptyState() {
  document.getElementById("no-recipients").style.display = "none";
}

function showRecipientsList() {
  document.getElementById("recipients-list").style.display = "block";
}

function hideRecipientsList() {
  document.getElementById("recipients-list").style.display = "none";
}

// ============================================
// Secure Send Toggle
// ============================================

function initializeSecureToggle() {
  const button = document.getElementById("secureToggle") as HTMLButtonElement;
  const icon = button.querySelector(".ig-toggle-icon");
  const label = button.querySelector(".ig-toggle-label");
  
  // Set initial state (not secure)
  button.setAttribute("aria-pressed", "false");
  icon.innerHTML = '🔓'; // Unlocked icon
  label.textContent = 'Not Secure';
}

function toggleSecureSend() {
  const button = document.getElementById("secureToggle") as HTMLButtonElement;
  const icon = button.querySelector(".ig-toggle-icon");
  const label = button.querySelector(".ig-toggle-label");
  
  const isPressed = button.getAttribute("aria-pressed") === "true";
  const newState = !isPressed;
  
  button.setAttribute("aria-pressed", newState.toString());
  
  if (newState) {
    // Secure mode ON
    icon.innerHTML = '🔒'; // Lock icon
    label.textContent = 'Secure Send';
  } else {
    // Secure mode OFF
    icon.innerHTML = '🔓'; // Unlocked icon
    label.textContent = 'Not Secure';
  }
  
  // Future: Add header logic here
  console.log('Secure send toggled:', newState);
}

// ============================================
// Version Display
// ============================================

function displayVersionInfo() {
  const versionEl = document.getElementById("app-version");
  const gitCommitEl = document.getElementById("git-commit");
  const gitCommitContainer = document.getElementById("git-commit-container");
  
  // Display version
  versionEl.textContent = APP_VERSION;
  
  // Display git commit if available
  if (GIT_COMMIT && GIT_COMMIT !== "dev") {
    gitCommitEl.textContent = GIT_COMMIT.substring(0, 7); // Short hash
    gitCommitContainer.style.display = "flex";
  }
}