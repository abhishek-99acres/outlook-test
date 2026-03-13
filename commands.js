// =============================================================================
//  commands.js  – Attachment Categorizer (Prefix-Rename Flow)
//
//  Events:
//    OnNewMessageCompose            → check pre-filled recipients
//    OnMessageAttachmentsChanged    → prompt user to categorize when trigger recipient present
//    OnMessageRecipientsChanged     → re-evaluate when To/CC changes
// =============================================================================
"use strict";

// ── CONFIGURATION ─────────────────────────────────────────────────────────────

const TRIGGER_RECIPIENT_MAP = {
  "abhishek.a3@99acres.com": "Abhishek Anand",
  "sonia.m@99acres.com":     "Sonia M",
  "finance@contoso.com":     "Finance",
  "legal@contoso.com":       "Legal",
  "hr@contoso.com":          "HR",
  "compliance@contoso.com":  "Compliance",
  "audit@contoso.com":       "Audit",
};

const NOTIF_KEY_ACTION   = "attachCatAction";
const NOTIF_KEY_NO_RULE  = "attachCatNoRule";
const TASKPANE_URL_RESID = "Taskpane.Url";

// ── ENTRY POINTS ──────────────────────────────────────────────────────────────

async function onNewMessageComposeHandler(event) {
  try   { await checkAndNotify(); }
  catch (e) { console.error("[AttachCat]", e); }
  finally   { event.completed(); }
}

async function onMessageAttachmentsChangedHandler(event) {
  try   { await checkAndNotify(); }
  catch (e) { console.error("[AttachCat]", e); }
  finally   { event.completed(); }
}

async function onMessageRecipientsChangedHandler(event) {
  try   { await checkAndNotify(); }
  catch (e) { console.error("[AttachCat]", e); }
  finally   { event.completed(); }
}

// ── CORE ──────────────────────────────────────────────────────────────────────

async function checkAndNotify() {
  const item = Office.context.mailbox.item;

  const [toR, ccR] = await Promise.all([
    getRecipientsAsync(item.to),
    getRecipientsAsync(item.cc),
  ]);

  const allRecipients  = [...toR, ...ccR];
  const matchedLabels  = getMatchedLabels(allRecipients);

  // No trigger recipient → clear and exit
  if (matchedLabels.length === 0) {
    item.notificationMessages.removeAsync(NOTIF_KEY_ACTION);
    item.notificationMessages.replaceAsync(NOTIF_KEY_NO_RULE, {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "No categorisation rules apply to current recipients.",
      icon: "Icon.16x16",
      persistent: false,
    });
    return;
  }

  // Check attachments
  const attachments = item.attachments || [];

  if (attachments.length === 0) {
    item.notificationMessages.removeAsync(NOTIF_KEY_NO_RULE);
    item.notificationMessages.replaceAsync(NOTIF_KEY_ACTION, {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `Sending to ${matchedLabels.join(", ")} — attach files and open the categorizer to label them.`,
      icon: "Icon.16x16",
      persistent: true,
    });
    return;
  }

  // Attachments present → prompt to categorize
  const uncategorized = attachments.filter(a => !a.name.includes("_"));
  const msg = uncategorized.length > 0
    ? `${uncategorized.length} attachment(s) need categorization. Open "View Categories" to label them.`
    : `All ${attachments.length} attachment(s) categorized. Open "View Categories" to review.`;

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
  return new Promise(resolve => {
    if (!field || typeof field.getAsync !== "function") return resolve([]);
    field.getAsync({}, result =>
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || [] : [])
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
if (typeof Office !== "undefined") {
  Office.actions.associate("onNewMessageComposeHandler",         onNewMessageComposeHandler);
  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler",  onMessageRecipientsChangedHandler);
}
