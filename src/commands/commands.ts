/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
export function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

function addSecureHeader(event) {
  // Get the current email item
  const item = Office.context.mailbox.item;
  
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
      
      // Signal that the function is complete
      event.completed();
    }
  );
}

// Register the function
Office.actions.associate("addSecureHeader", addSecureHeader);