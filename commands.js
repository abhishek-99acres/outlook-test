"use strict";

// ── CONFIG ────────────────────────────────────────────────────────────────────

const TRIGGER_RECIPIENT_MAP = {
  "abhishek.a3@99acres.com":  "Abhishek Anand",
  "sonia.m@99acres.com":      "Sonia M",
  "coder.abhi02@gmail.com":   "Abhishek Kumar",
  "rkd02122@gmail.com":       "Abhishek Kumar",
  "finance@contoso.com":      "Finance",
  "legal@contoso.com":        "Legal",
  "hr@contoso.com":           "HR",
  "compliance@contoso.com":   "Compliance",
  "audit@contoso.com":        "Audit",
};

const KNOWN_PREFIXES = [
  "Legal_", "Finance_", "HR_", "Compliance_", "Contract_",
  "Invoice_", "Report_", "Presentation_", "Reference_", "General_",
];

// ── SEND HANDLER ──────────────────────────────────────────────────────────────
// Synchronous. No await. No API calls. Done in <1ms.
// Checks if filename starts with a known prefix (set by taskpane on Apply).

function onMessageSendHandler(event) {
  try {
    var attachments = Office.context.mailbox.item.attachments || [];

    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    var bad = attachments.filter(function (a) {
      return !KNOWN_PREFIXES.some(function (p) { return a.name.startsWith(p); });
    });

    if (bad.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    event.completed({
      allowEvent: false,
      errorMessage:
        bad.length + " attachment(s) not categorized:\n" +
        bad.map(function (a) { return "  \u2022 " + a.name; }).join("\n") +
        "\n\nOpen 'View Categories' in the ribbon to label them.",
    });

  } catch (e) {
    console.error("[AttachCat] send error:", e);
    event.completed({ allowEvent: true });
  }
}

// ── COMPOSE EVENTS ────────────────────────────────────────────────────────────

function onNewMessageComposeHandler(event) {
  try { updateNotification(); } catch (e) {}
  event.completed();
}

function onMessageAttachmentsChangedHandler(event) {
  try {
    updateNotification();
    var attachments = Office.context.mailbox.item.attachments || [];
    var hasUncategorized = attachments.some(function (a) {
      return !KNOWN_PREFIXES.some(function (p) { return a.name.startsWith(p); });
    });
    if (hasUncategorized && Office.addin && Office.addin.showAsTaskpane) {
      Office.addin.showAsTaskpane();
    }
  } catch (e) {}
  event.completed();
}

function onMessageRecipientsChangedHandler(event) {
  try { updateNotification(); } catch (e) {}
  event.completed();
}

// ── NOTIFICATION BAR ──────────────────────────────────────────────────────────

function updateNotification() {
  var attachments = Office.context.mailbox.item.attachments || [];
  var bad = attachments.filter(function (a) {
    return !KNOWN_PREFIXES.some(function (p) { return a.name.startsWith(p); });
  });
  var msg = attachments.length === 0
    ? "Attach a file — the categorizer will open automatically."
    : bad.length > 0
      ? "\u26a0\ufe0f " + bad.length + " attachment(s) need a category before sending."
      : "\u2713 All " + attachments.length + " attachment(s) categorized. Ready to send.";

  Office.context.mailbox.item.notificationMessages.replaceAsync("attachCat", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: "Icon.16x16",
    persistent: true,
  });
}

// ── REGISTER ──────────────────────────────────────────────────────────────────

if (typeof Office !== "undefined") {
  Office.actions.associate("onNewMessageComposeHandler",         onNewMessageComposeHandler);
  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler",  onMessageRecipientsChangedHandler);
  Office.actions.associate("onMessageSendHandler",               onMessageSendHandler);
}
