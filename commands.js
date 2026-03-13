// =============================================================================
//  commands.js  – Attachment Categorizer
//  Runtime is kept alive with lifetime="long" in manifest.
//  onMessageSendHandler is therefore ALWAYS called into an already-loaded
//  runtime — no cold-start, no timeout.
// =============================================================================
"use strict";

// ── CONFIGURATION ─────────────────────────────────────────────────────────────

const TRIGGER_RECIPIENT_MAP = {
  "abhishek.a3@99acres.com": "Abhishek Anand",
  "sonia.m@99acres.com": "Sonia M",
  "coder.abhi02@gmail.com": "Abhishek Kumar",
  "rkd02122@gmail.com": "Abhishek Kumar",
  "finance@contoso.com": "Finance",
  "legal@contoso.com": "Legal",
  "hr@contoso.com": "HR",
  "compliance@contoso.com": "Compliance",
  "audit@contoso.com": "Audit",
};

// Must exactly match CATEGORIES in taskpane.html
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

const NOTIF_KEY_ACTION = "attachCatAction";
const NOTIF_KEY_NO_RULE = "attachCatNoRule";

// ── SEND GUARD ────────────────────────────────────────────────────────────────
// Synchronous — completes in <1ms.
// Safe because lifetime="long" means this runtime is already loaded.
// =============================================================================
function onMessageSendHandler(event) {
  try {
    const attachments = Office.context.mailbox.item.attachments || [];

    // No attachments → allow
    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    // Find files without a category prefix
    const uncategorized = attachments.filter(
      (att) => !KNOWN_PREFIXES.some((p) => att.name.startsWith(p)),
    );

    if (uncategorized.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    // Block with our custom message
    event.completed({
      allowEvent: false,
      errorMessage:
        uncategorized.length +
        " attachment(s) not categorized:\n\n" +
        uncategorized.map((a) => "• " + a.name).join("\n") +
        "\n\nOpen 'View Categories' in the ribbon to label each file.",
    });
  } catch (err) {
    // Safety net — on any unexpected error still block so user is aware
    event.completed({
      allowEvent: false,
      errorMessage:
        "Attachment check failed: " +
        (err.message || err) +
        "\nPlease verify attachments are categorized before sending.",
    });
  }
}

// ── COMPOSE EVENTS ────────────────────────────────────────────────────────────

async function onNewMessageComposeHandler(event) {
  try {
    await checkAndNotify();
  } catch (e) {
    console.error("[AttachCat]", e);
  } finally {
    event.completed();
  }
}

async function onMessageAttachmentsChangedHandler(event) {
  try {
    await checkAndNotify();
  } catch (e) {
    console.error("[AttachCat]", e);
  } finally {
    event.completed();
  }
}

async function onMessageRecipientsChangedHandler(event) {
  try {
    await checkAndNotify();
  } catch (e) {
    console.error("[AttachCat]", e);
  } finally {
    event.completed();
  }
}

async function checkAndNotify() {
  const item = Office.context.mailbox.item;

  const [toR, ccR] = await Promise.all([
    getRecipientsAsync(item.to),
    getRecipientsAsync(item.cc),
  ]);

  const matchedLabels = getMatchedLabels([...toR, ...ccR]);

  if (matchedLabels.length === 0) {
    item.notificationMessages.removeAsync(NOTIF_KEY_ACTION);
    item.notificationMessages.replaceAsync(NOTIF_KEY_NO_RULE, {
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
      message: "No categorisation rules apply to current recipients.",
      icon: "Icon.16x16",
      persistent: false,
    });
    return;
  }

  const attachments = item.attachments || [];
  const uncategorized = attachments.filter(
    (a) => !KNOWN_PREFIXES.some((p) => a.name.startsWith(p)),
  );

  const msg =
    attachments.length === 0
      ? "Sending to " +
        matchedLabels.join(", ") +
        ' — open "View Categories" to label attachments before sending.'
      : uncategorized.length > 0
        ? "⚠️ " +
          uncategorized.length +
          " of " +
          attachments.length +
          ' attachment(s) not yet labeled. Open "View Categories".'
        : "✓ All " +
          attachments.length +
          " attachment(s) categorized. Ready to send.";

  item.notificationMessages.removeAsync(NOTIF_KEY_NO_RULE);
  // item.notificationMessages.replaceAsync(NOTIF_KEY_ACTION, {
  //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //   message: msg,
  //   icon: "Icon.16x16",
  //   persistent: true,
  // });

  item.notificationMessages.replaceAsync(NOTIF_KEY_ACTION, {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: "Icon.16x16",
    persistent: true,
    actions: [
      {
        actionText: "Open Attachment Categorizer",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
        commandId: "msgComposeOpenPaneButton",
      },
    ],
  });
}

// ── HELPERS ───────────────────────────────────────────────────────────────────

function getRecipientsAsync(field) {
  return new Promise((resolve) => {
    if (!field || typeof field.getAsync !== "function") return resolve([]);
    field.getAsync({}, (r) =>
      resolve(
        r.status === Office.AsyncResultStatus.Succeeded ? r.value || [] : [],
      ),
    );
  });
}

function getMatchedLabels(recipients) {
  const matched = new Set();
  for (const r of recipients) {
    const email = (r.emailAddress || "").toLowerCase().trim();
    if (TRIGGER_RECIPIENT_MAP[email]) matched.add(TRIGGER_RECIPIENT_MAP[email]);
  }
  return [...matched];
}

// ── REGISTER ──────────────────────────────────────────────────────────────────
Office.onReady(function () {
  // Register all handlers after Office is ready
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
});
