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


function toggleSecureHeader(event) {
  const item = Office.context.mailbox.item;
  
  // Load custom properties to check current state
  item.loadCustomPropertiesAsync(function(result) {
    const customProps = result.value;
    const isEnabled = customProps.get("SecureHeaderEnabled") === "true";
    
    if (isEnabled) {
      // Currently enabled, so DISABLE it
      item.internetHeaders.removeAsync(["X-Secure-Send"], function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          customProps.set("SecureHeaderEnabled", "false");
          customProps.saveAsync(function() {
            showNotification("Header Disabled", "X-Secure-Send header removed");
            item.notificationMessages.removeAsync("secureHeaderInfo");
            event.completed();
          });
        } else {
          showNotification("Error", asyncResult.error.message, "errorMessage");
          event.completed();
        }
      });
    } else {
      // Currently disabled, so ENABLE it
      item.internetHeaders.setAsync({ "X-Secure-Send": "1" }, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          customProps.set("SecureHeaderEnabled", "true");
          customProps.saveAsync(function() {
            showNotification("Header Enabled", "X-Secure-Send header will be added");

            // Show persistent info bar
            item.notificationMessages.addAsync(
              "secureHeaderInfo",
              {
                type: "informationalMessage",
                message: "✓ Secure Send is ENABLED for this email",
                icon: "Icon.16x16",
                persistent: true
              }
            );

            event.completed();
          });
        } else {
          showNotification("Error", asyncResult.error.message, "errorMessage");
          event.completed();
        }
      });
    }
  });
}

function showNotification(title, message, type = "informationalMessage") {
  Office.context.mailbox.item.notificationMessages.addAsync(
    "headerNotification",
    {
      type: type,
      message: message,
      icon: "Icon.16x16",
      persistent: true
    }
  );
}

// Register the toggle function
Office.actions.associate("toggleSecureHeader", toggleSecureHeader);