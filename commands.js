// =============================================================================
//  commands.js  – Attachment Categorizer
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

// ⚠️ MUST match the CATEGORIES list in taskpane.html exactly
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

// =============================================================================
//  SEND GUARD (Blocks send & opens Taskpane)
// =============================================================================
async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    // 1. Check if trigger recipients are present
    const [toR, ccR] = await Promise.all([
      getRecipientsAsync(item.to),
      getRecipientsAsync(item.cc),
    ]);

    const matchedLabels = getMatchedLabels([...toR, ...ccR]);

    // If no trigger recipients, allow sending
    if (matchedLabels.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    // 2. Check attachments accurately
    const attachments = await getAttachmentsAsync(item);

    if (attachments.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    // 3. Find uncategorized attachments
    const uncategorized = attachments.filter(
      (att) => !KNOWN_PREFIXES.some((p) => att.name.startsWith(p)),
    );

    if (uncategorized.length === 0) {
      // All categorized, allow sending
      event.completed({ allowEvent: true });
      return;
    }

    // 4. Block send AND open the taskpane
    event.completed({
      allowEvent: false,
      errorMessage: `⚠️ You have ${uncategorized.length} attachment(s) that need to be categorized before sending.`,
      cancelLabel: "Categorize Now", // The text on the button Outlook shows
      commandId: "msgComposeOpenPaneButton", // Matches the Control ID in manifest.xml to open the UI
    });
  } catch (error) {
    console.error("[AttachCat] Error during onMessageSend:", error);
    // FAIL-SAFE: Always allow send if the script crashes so the user is not stuck
    event.completed({ allowEvent: true });
  }
}

// =============================================================================
//  COMPOSE / ATTACHMENT / RECIPIENT EVENTS
// =============================================================================

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

  const attachments = await getAttachmentsAsync(item);
  const uncategorized = attachments.filter(
    (a) => !KNOWN_PREFIXES.some((p) => a.name.startsWith(p)),
  );

  const msg =
    attachments.length === 0
      ? "Sending to " +
        matchedLabels.join(", ") +
        ' — attach files and open "View Categories" to label them.'
      : uncategorized.length > 0
        ? "⚠️ " +
          uncategorized.length +
          " of " +
          attachments.length +
          ' attachment(s) still need a category. Open "View Categories".'
        : "✓ All " + attachments.length + " attachment(s) categorized.";

  item.notificationMessages.removeAsync(NOTIF_KEY_NO_RULE);
  item.notificationMessages.replaceAsync(NOTIF_KEY_ACTION, {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msg,
    icon: "Icon.16x16",
    persistent: true,
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

function getAttachmentsAsync(item) {
  return new Promise((resolve) => {
    if (typeof item.getAttachmentsAsync === "function") {
      item.getAttachmentsAsync({}, (r) => {
        resolve(
          r.status === Office.AsyncResultStatus.Succeeded ? r.value || [] : [],
        );
      });
    } else {
      // Fallback for older versions
      resolve(item.attachments || []);
    }
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
if (typeof Office !== "undefined") {
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
}
