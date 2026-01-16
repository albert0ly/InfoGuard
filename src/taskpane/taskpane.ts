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

let currentDialog: Office.Dialog | null = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getProfileButton").onclick = run;
    document.getElementById("loadRecipientsButton").onclick = loadRecipients;
    document.getElementById("secureToggle").onclick = toggleSecureSend;
    
    // Dialog button handlers
    document.getElementById("openDialogOption1").onclick = openDialogOption1;
    document.getElementById("openDialogOption2").onclick = openDialogOption2;
    
    // Initialize secure toggle button icon
    initializeSecureToggle();
    
    // Display version info
    displayVersionInfo();
    
    // Auto-load recipients when task pane opens (if composing)
    if (Office.context.mailbox.item) {
      loadRecipients();
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
// Dialog Options
// ============================================

/**
 * Option 1: Office Dialog API (Centered on entire Outlook window)
 * This uses displayDialogAsync with strategies to minimize popup warnings
 */
function openDialogOption1() {
  // Close existing dialog if any
  if (currentDialog) {
    currentDialog.close();
    currentDialog = null;
  }

  // IMPORTANT: Dialog URL must be absolute and on same domain as your add-in
  // For development: use your dev server URL
  // For production: use your deployed domain
  // If placed in assets folder, use: window.location.origin + "/assets/dialog.html"
  const dialogUrl = window.location.origin + "/dialog.html";
  
  const dialogOptions: Office.DialogOptions = {
    height: 60, // Percentage of screen height
    width: 50,  // Percentage of screen width
    displayInIframe: true, // TRUE to minimize popup warnings
    promptBeforeOpen: false
  };

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    dialogOptions,
    (asyncResult: Office.AsyncResult<Office.Dialog>) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Dialog failed to open:", asyncResult.error.message);
        
        // Show error using Office notification instead of alert
        let errorMessage = "Failed to open dialog: " + asyncResult.error.message;
        
        if (asyncResult.error.code === 12004) {
          errorMessage = "Please allow popups for this add-in to open dialogs.";
        } else if (asyncResult.error.code === 12005) {
          errorMessage = "A dialog is already open.";
        }
        
        // Display error in the validation error area
        showError(errorMessage);
      } else {
        currentDialog = asyncResult.value;
        
        // Listen for messages from dialog
        currentDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          processDialogMessage
        );
        
        // Listen for dialog events (close, errors)
        currentDialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          processDialogEvent
        );
      }
    }
  );
}

/**
 * Option 2: Custom Modal (Only covers the taskpane)
 * This doesn't center on entire Outlook, but never shows popup warnings
 */
function openDialogOption2() {
  const overlay = document.getElementById('customModalOverlay');
  if (overlay) {
    overlay.classList.add('active');
    
    // Focus first input after a brief delay
    setTimeout(() => {
      const firstInput = document.getElementById('modalNameInput') as HTMLInputElement;
      if (firstInput) {
        firstInput.focus();
      }
    }, 100);
  }
}

/**
 * Handle messages received from the Office Dialog (Option 1)
 */
function processDialogMessage(arg: { message: string; origin: string | undefined }) {
  try {
    const messageFromDialog = JSON.parse(arg.message);
    console.log("Received from dialog:", messageFromDialog);
    
    // Process different actions
    switch (messageFromDialog.action) {
      case "submit":
        handleDialogSubmit(messageFromDialog.data);
        break;
      case "cancel":
        console.log("Dialog cancelled");
        break;
    }
    
    // Close dialog after processing
    if (currentDialog) {
      currentDialog.close();
      currentDialog = null;
    }
  } catch (error) {
    console.error("Error processing dialog message:", error);
  }
}

/**
 * Handle dialog events (user closed dialog, navigation errors, etc.)
 */
function processDialogEvent(arg: { error: number; type?: string }) {
  console.log("Dialog event:", arg);
  
  switch (arg.error) {
    case 12002:
      console.log("User closed dialog");
      break;
    case 12003:
      console.error("Dialog navigation failed");
      break;
    case 12006:
      console.error("Dialog sent too many messages");
      break;
  }
  
  currentDialog = null;
}

/**
 * Handle form submission from dialog
 */
function handleDialogSubmit(data: any) {
  console.log("Processing dialog submission:", data);
  
  // Example: Insert data into email body
  if (Office.context.mailbox && Office.context.mailbox.item) {
    const formattedData = `Name: ${data.name}\nEmail: ${data.email}\nMessage: ${data.message}`;
    
    Office.context.mailbox.item.body.setSelectedDataAsync(
      formattedData,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Data inserted into email");
          
          // Show success notification
          if (Office.context.mailbox && Office.context.mailbox.item) {
            Office.context.mailbox.item.notificationMessages.addAsync(
              "dialogDataInserted",
              {
                type: "informationalMessage",
                message: "Dialog data inserted successfully",
                icon: "Icon.16x16",
                persistent: false
              }
            );
          }
        } else {
          console.error("Failed to insert data:", result.error);
          showError("Failed to insert dialog data into email.");
        }
      }
    );
  }
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
    
    row.innerHTML = `
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
      <div class="ig-list-cell ig-cell-actions">
        <button class="ig-actions-btn" title="More">...</button>
      </div>
    `;
    
    tbody.appendChild(row);
    
    // Add event listener for validation
    const input = document.getElementById(mobileInputId) as HTMLInputElement;
    input.addEventListener('input', (e) => validateMobileInput(e.target as HTMLInputElement));
    input.addEventListener('blur', (e) => validateMobileInput(e.target as HTMLInputElement));
  });

  showRecipientsList();
  clearError();
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
  if (!value) {
    input.classList.remove('invalid');
    return;
  }
  const isValid = validateIsraeliMobile(value);
  if (isValid) {
    input.classList.remove('invalid');
  } else {
    input.classList.add('invalid');
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
  const wantOn = !isPressed;

  if (wantOn) {
    const ok = validateAllRecipients();
    if (!ok) {
      button.setAttribute("aria-pressed", "false");
      icon.innerHTML = '🔓';
      label.textContent = 'Not Secure';
      showError('Cannot enable Secure Send due to validation errors.');
      return;
    }
  } else {
    clearError();
  }
  
  button.setAttribute("aria-pressed", wantOn.toString());
  
  if (wantOn) {
    addSecureHeader();
    icon.innerHTML = '🔒';
    label.textContent = 'Secure Send';
  } else {
    removeSecureHeader();
    icon.innerHTML = '🔓';
    label.textContent = 'Not Secure';
  }
  
  console.log('Secure send toggled:', wantOn);
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

// ============================================
// Validation and Error Display Functions
// ============================================

/**
 * Validates that all recipients have a mobile number
 * @returns true if all recipients have valid mobile numbers, false otherwise
 */
function validateAllRecipients(): boolean {
  const inputs = document.querySelectorAll('.ig-mobile-input') as NodeListOf<HTMLInputElement>;
  
  if (inputs.length === 0) {
    showError('No recipients found. Please add recipients to enable secure send.');
    return false;
  }
  
  const invalidRecipients: string[] = [];
  
  inputs.forEach(input => {
    const email = input.getAttribute('data-email');
    const mobile = input.value.trim();
    
    if (!mobile) {
      invalidRecipients.push(email);
    } else if (!validateIsraeliMobile(mobile)) {
      invalidRecipients.push(email);
    }
  });
  
  if (invalidRecipients.length > 0) {
    const recipientList = invalidRecipients.join(', ');
    const message = invalidRecipients.length === 1
      ? `Missing or invalid mobile number for: ${recipientList}`
      : `Missing or invalid mobile numbers for ${invalidRecipients.length} recipients: ${recipientList}`;
    
    showError(message);
    return false;
  }
  
  return true;
}

/**
 * Displays an error message in the validation error container
 * @param message - The error message to display
 */
function showError(message: string): void {
  const errorEl = document.getElementById('validation-error');
  if (errorEl) {
    errorEl.textContent = message;
    errorEl.style.display = 'block';
    
    // Scroll error into view
    errorEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

/**
 * Clears the error message from the validation error container
 */
function clearError(): void {
  const errorEl = document.getElementById('validation-error');
  if (errorEl) {
    errorEl.textContent = '';
    errorEl.style.display = 'none';
  }
}

function addSecureHeader() {
  // Get the current email item
  const item = Office.context.mailbox.item;
  
  item.notificationMessages.removeAsync("headerNotification");

  // Add custom internet header
  item.internetHeaders.setAsync(
    { "X-Secure-Send": "1" },
    function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // Show success notification
        Office.context.mailbox.item.notificationMessages.addAsync(
          "headerAdded",
          {
            type: "informationalMessage",
            message: "X-Secure-Send header added successfully",
            icon: "Icon.16x16",
            persistent: false
          }
        );
      } else {
        // Show error
        Office.context.mailbox.item.notificationMessages.addAsync(
          "headerError",
          {
            type: "errorMessage",
            message: "Failed to add header: " + asyncResult.error.message
          }
        );
      }
    }
  );
}

function removeSecureHeader() {
  // Get the current email item
  const item = Office.context.mailbox.item;

  item.notificationMessages.removeAsync("headerNotification");

  // Remove custom internet header
  item.internetHeaders.removeAsync(
    ["X-Secure-Send"],
    function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // Show success notification
        Office.context.mailbox.item.notificationMessages.addAsync(
          "headerRemoved",
          {
            type: "informationalMessage",
            message: "X-Secure-Send header removed successfully",
            icon: "Icon.16x16",
            persistent: false
          }
        );
      } else {
        // Show error
        Office.context.mailbox.item.notificationMessages.addAsync(
          "headerError",
          {
            type: "errorMessage",
            message: "Failed to remove header: " + asyncResult.error.message
          }
        );
      }
    }
  );
}