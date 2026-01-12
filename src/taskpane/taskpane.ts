/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { getUserData, getGraphToken } from "../helpers/sso-helper";

interface Recipient {
  no: number;
  email: string;
  mobile: string | null;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getProfileButton").onclick = run;
    document.getElementById("loadRecipientsButton").onclick = loadRecipients;
    
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

    // Fetch mobile numbers for each recipient
    const recipientsWithMobile = await Promise.all(
      recipients.map(async (recipient, index) => {
        const mobile = await getMobileNumber(recipient.email, token);
        return {
          no: index + 1,
          email: recipient.email,
          mobile: mobile
        };
      })
    );

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
    
    const mobileDisplay = recipient.mobile 
      ? `<span class="ig-mobile-badge"><i class="ms-Icon ms-Icon--Phone"></i>${recipient.mobile}</span>`
      : `<span style="color: #a19f9d;">N/A</span>`;
    
    row.innerHTML = `
      <div class="ig-list-cell ig-cell-no">${recipient.no}</div>
      <div class="ig-list-cell ig-cell-email" title="${recipient.email}">${recipient.email}</div>
      <div class="ig-list-cell ig-cell-mobile">${mobileDisplay}</div>
    `;
    
    tbody.appendChild(row);
  });

  showRecipientsList();
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