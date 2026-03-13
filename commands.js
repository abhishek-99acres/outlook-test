"use strict";

const KNOWN_PREFIXES = [
  "Legal_", "Finance_", "HR_", "Compliance_",
  "Contract_", "Invoice_", "Report_",
  "Presentation_", "Reference_", "General_",
];

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

// ── SEND GUARD — synchronous, zero async calls ───────────────────────────────
function onMessageSendHandler(event) {
  try {
    const attachments = Office.context.mailbox.item.attachments || [];

    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    const uncategorized = attachments.filter(
      att => !KNOWN_PREFIXES.some(p => att.name.startsWith(p))
    );

    if (uncategorized.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    event.completed({
      allowEvent: false,
      errorMessage:
        uncategorized.length + " attachment(s) not categorized:\n" +
        uncategorized.map(a => "• " + a.name).join("\n") +
        "\n\nOpen 'View Categories' in the ribbon and label each file.",
    });
  } catch (e) {
    event.completed({ allowEvent: true });
  }
}

// ── COMPOSE OPEN ─────────────────────────────────────────────────────────────
function onNewMessageComposeHandler(event) {
  updateNotification();
  event.completed();
}

// ── ATTACHMENT ADDED — auto-open taskpane immediately ────────────────────────
function onMessageAttachmentsChangedHandler(event) {
  updateNotification();

  try {
    const attachments = Office.context.mailbox.item.attachments || [];
    const hasUncategorized = attachments.some(
      att => !KNOWN_PREFIXES.some(p => att.name.startsWith(p))
    );

    // Auto-pop the taskpane the moment an uncategorized file is attached
    if (hasUncategorized && Office.addin && typeof Office.addin.showAsTaskpane === "function") {
      Office.addin.showAsTaskpane();
    }
  } catch (e) {
    console.error("[AttachCat] showAsTaskpane error:", e);
  }

  event.completed();
}

// ── RECIPIENTS CHANGED ───────────────────────────────────────────────────────
function onMessageRecipientsChangedHandler(event) {
  updateNotification();
  event.completed();
}

// ── NOTIFICATION BAR ─────────────────────────────────────────────────────────
function updateNotification() {
  try {
    const item = Office.context.mailbox.item;
    const attachments = item.attachments || [];
    const uncategorized = attachments.filter(
      a => !KNOWN_PREFIXES.some(p => a.name.startsWith(p))
    );

    const msg = attachments.length === 0
      ? "Attach files — taskpane will open automatically to categorize them."
      : uncategorized.length > 0
        ? "⚠️ " + uncategorized.length + " attachment(s) need a category label before sending."
        : "✓ All attachments categorized. Ready to send.";

    item.notificationMessages.replaceAsync("attachCat", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: msg,
      icon: "Icon.16x16",
      persistent: true,
    });
  } catch (e) {
    console.error("[AttachCat] updateNotification:", e);
  }
}

// ── REGISTER ─────────────────────────────────────────────────────────────────
Office.onReady(function () {
  Office.actions.associate("onNewMessageComposeHandler",         onNewMessageComposeHandler);
  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler",  onMessageRecipientsChangedHandler);
  Office.actions.associate("onMessageSendHandler",               onMessageSendHandler);
});
