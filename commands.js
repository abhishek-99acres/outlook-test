"use strict";

/* =============================================================================
   Attachment Categorizer — commands.js

   CRITICAL: For event-based activation, ALL Office.actions.associate() calls
   MUST be at the TOP LEVEL of the script — NOT inside Office.onReady().
   Office.onReady() is for taskpanes only. Putting handlers inside it causes
   the 5-second timeout because Office waits for the association before the
   callback ever runs.
============================================================================= */

const KNOWN_PREFIXES = [
  "Legal_",
  "Finance_",
  "HR_",
  "Compliance_",
  "Contract_",
  "Invoice_",
  "Report_",
  "Presentation_",
  "Reference_",
  "General_",
];

// ---------------------------------------------------------------------------
// OnMessageSend — synchronous, zero API calls, completes in <1ms
// ---------------------------------------------------------------------------
function onMessageSendHandler(event) {
  try {
    const attachments = Office.context.mailbox.item.attachments || [];

    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    const bad = attachments.filter(function (a) {
      return !KNOWN_PREFIXES.some(function (p) {
        return a.name.startsWith(p);
      });
    });

    if (bad.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    event.completed({
      allowEvent: false,
      errorMessage:
        bad.length +
        " attachment(s) not categorized:\n" +
        bad
          .map(function (a) {
            return "  \u2022 " + a.name;
          })
          .join("\n") +
        "\n\nOpen 'View Categories' in the ribbon to label them.",
    });
  } catch (e) {
    console.error("[AttachCat] send error:", e);
    event.completed({ allowEvent: true });
  }
}

// ---------------------------------------------------------------------------
// OnNewMessageCompose
// ---------------------------------------------------------------------------
function onNewMessageComposeHandler(event) {
  try {
    showNotification();
  } catch (e) {}
  event.completed();
}

// ---------------------------------------------------------------------------
// OnMessageAttachmentsChanged — auto-open taskpane
// ---------------------------------------------------------------------------
function onMessageAttachmentsChangedHandler(event) {
  try {
    showNotification();
    var attachments = Office.context.mailbox.item.attachments || [];
    var hasUncategorized = attachments.some(function (a) {
      return !KNOWN_PREFIXES.some(function (p) {
        return a.name.startsWith(p);
      });
    });
    if (hasUncategorized && Office.addin && Office.addin.showAsTaskpane) {
      Office.addin.showAsTaskpane();
    }
  } catch (e) {
    console.error("[AttachCat] attachment change error:", e);
  }
  event.completed();
}

// ---------------------------------------------------------------------------
// OnMessageRecipientsChanged
// ---------------------------------------------------------------------------
function onMessageRecipientsChangedHandler(event) {
  try {
    showNotification();
  } catch (e) {}
  event.completed();
}

// ---------------------------------------------------------------------------
// Notification bar
// ---------------------------------------------------------------------------
function showNotification() {
  var attachments = Office.context.mailbox.item.attachments || [];
  var bad = attachments.filter(function (a) {
    return !KNOWN_PREFIXES.some(function (p) {
      return a.name.startsWith(p);
    });
  });
  var msg =
    attachments.length === 0
      ? "Attach a file — the categorizer will open automatically."
      : bad.length > 0
        ? "\u26a0\ufe0f " +
          bad.length +
          " attachment(s) need a category label before sending."
        : "\u2713 All " +
          attachments.length +
          " attachment(s) categorized. Ready to send.";

  Office.context.mailbox.item.notificationMessages.replaceAsync("attachCat", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: "Icon.16x16",
    persistent: true,
  });
}

// ---------------------------------------------------------------------------
// REGISTER AT TOP LEVEL — this is mandatory for event-based activation.
// Do NOT wrap in Office.onReady() — that delays registration past the timeout.
// ---------------------------------------------------------------------------
Office.actions.associate(
  "onNewMessageComposeHandler",
  onNewMessageComposeHandler,
);
Office.actions.associate(
  "onMessageAttachmentsChangedHandler",
  onMessageAttachmentsChangedHandler,
);
Office.actions.associate(
  "onMessageRecipientsChangedHandler",
  onMessageRecipientsChangedHandler,
);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
