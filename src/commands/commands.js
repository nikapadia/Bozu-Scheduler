/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
  console.log("Office is ready!");
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Settings.80x80",
    persistent: true,
  };

  // console.log("Hi there");
  function getStartTime(callback) {
    Office.context.mailbox.item.start.getAsync(function (result) {
      var value = result.value.toISOString();
      callback(value);
    });
  }
  function getEndTime(callback) {
    Office.context.mailbox.item.end.getAsync(function (result) {
      var value = result.value.toISOString();
      callback(value);
    });
  }
  function getSubject(callback) {
    Office.context.mailbox.item.subject.getAsync(function (result) {
      var value = result.value;
      callback(value);
    });
  }
  function getLocation(callback) {
    Office.context.mailbox.item.location.getAsync(function (result) {
      var value = result.value;
      callback(value);
    });
  }

  // Call the getValue function and pass a callback function to handle the value
  getStartTime(function (value) {
    console.log(value);
  });
  getEndTime(function (value) {
    console.log(value);
  });
  getSubject(function (value) {
    console.log(value);
  });
  getLocation(function (value) {
    console.log(value);
  });
  
  // Office.context.ui.displayDialogAsync('http://localhost:3000/');

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
