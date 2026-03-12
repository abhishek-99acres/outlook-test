// =============================================================================
//  commands.js  –  Attachment Categorizer: Event-based Activation Logic
//
//  Events handled:
//    onNewMessageComposeHandler         – fires when a new compose window opens
//    onMessageAttachmentsChangedHandler – fires when files are added or removed
//    onMessageRecipientsChangedHandler  – fires when To/CC recipients change
//
//  Flow:
//    1. Collect all To + CC recipients.
//    2. Check if any match the TRIGGER_RECIPIENT_MAP keys.
//    3. If a match is found, iterate every attachment and assign it a category
//       string derived from ATTACHMENT_CATEGORY_RULES.
//    4. Apply the resulting categories to the mail item and show a notification.
// =============================================================================

"use strict";

// ---------------------------------------------------------------------------
// ❶  CONFIGURATION – edit these two objects to match your business rules
// ---------------------------------------------------------------------------

/**
 * Map of email addresses (lower-cased) → human-readable label used in
 * category names and notifications.
 *
 * When ANY of these addresses appears in the To or CC field the
 * categorisation logic is triggered.
 *
 * Example entries:
 *   "finance@contoso.com"  → triggers Finance rules
 *   "legal@contoso.com"    → triggers Legal rules
 */
// const TRIGGER_RECIPIENT_MAP = {
//   "abhishek.a3@99Acres.com": "Abhishek Anand",
//   "finance@contoso.com": "Finance",
//   "legal@contoso.com": "Legal",
//   "hr@contoso.com": "HR",
//   "compliance@contoso.com": "Compliance",
//   "audit@contoso.com": "Audit",
// };

const TRIGGER_RECIPIENT_MAP = {
  "abhishek.a3@99acres.com": "Abhishek Anand", // ← lowercase 'a' in 99acres
  "finance@contoso.com": "Finance",
  "legal@contoso.com": "Legal",
  "hr@contoso.com": "HR",
  "compliance@contoso.com": "Compliance",
  "audit@contoso.com": "Audit",
};

/**
 * Rules that translate a file extension into a document-type label.
 * Keys must be lower-cased extensions WITHOUT the leading dot.
 * The special key "_default" is used when no extension matches.
 */
const ATTACHMENT_CATEGORY_RULES = {
  // Documents
  pdf: "PDF Document",
  doc: "Word Document",
  docx: "Word Document",
  // Spreadsheets
  xls: "Spreadsheet",
  xlsx: "Spreadsheet",
  csv: "Spreadsheet",
  // Presentations
  ppt: "Presentation",
  pptx: "Presentation",
  // Images
  jpg: "Image",
  jpeg: "Image",
  png: "Image",
  gif: "Image",
  bmp: "Image",
  svg: "Image",
  // Archives
  zip: "Archive",
  rar: "Archive",
  "7z": "Archive",
  tar: "Archive",
  gz: "Archive",
  // Code / data
  json: "Data File",
  xml: "Data File",
  txt: "Text File",
  // Fallback
  _default: "General Attachment",
};

// Outlook category colours available via Office JS
// See: https://learn.microsoft.com/en-us/javascript/api/outlook/office.mailboxenums.categorycolor
const RECIPIENT_CATEGORY_COLORS = {
  Finance: "Preset0", // Red
  Legal: "Preset1", // Orange
  HR: "Preset2", // Peach
  Compliance: "Preset3", // Yellow
  Audit: "Preset4", // Green
  _default: "Preset8", // Blue
};

// Notification keys (must be unique per notification)
const NOTIF_KEY_CATEGORIZED = "attachmentCategorized";
const NOTIF_KEY_NO_TRIGGER = "noTriggerRecipient";

// ---------------------------------------------------------------------------
// ❷  ENTRY POINTS  (registered in manifest.xml)
// ---------------------------------------------------------------------------

/**
 * OnNewMessageCompose – initialise; no attachments yet, but we can check
 * pre-filled recipients (e.g. when replying or using a template).
 */
async function onNewMessageComposeHandler(event) {
  try {
    await evaluateAndCategorize(event, "compose-open");
  } catch (err) {
    console.error("[AttachCat] onNewMessageCompose error:", err);
  } finally {
    event.completed();
  }
}

/**
 * OnMessageAttachmentsChanged – main trigger: re-run categorisation whenever
 * a file is attached or removed.
 */
async function onMessageAttachmentsChangedHandler(event) {
  try {
    const details = event.attachmentDetails; // added / removed info
    const action =
      details &&
      details.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added
        ? "added"
        : "removed";

    console.log(
      `[AttachCat] Attachment ${action}: ${details && details.attachmentName}`,
    );
    await evaluateAndCategorize(event, action);
  } catch (err) {
    console.error("[AttachCat] onMessageAttachmentsChanged error:", err);
  } finally {
    event.completed();
  }
}

/**
 * OnMessageRecipientsChanged – re-run when To/CC changes so categories stay
 * in sync even if recipients are added after attachments.
 */
async function onMessageRecipientsChangedHandler(event) {
  try {
    console.log("[AttachCat] Recipients changed – re-evaluating categories.");
    await evaluateAndCategorize(event, "recipients-change");
  } catch (err) {
    console.error("[AttachCat] onMessageRecipientsChanged error:", err);
  } finally {
    event.completed();
  }
}

// ---------------------------------------------------------------------------
// ❸  CORE LOGIC
// ---------------------------------------------------------------------------

/**
 * Main orchestration function.
 * 1. Reads all To + CC recipients.
 * 2. Finds matching trigger recipients.
 * 3. Reads current attachments.
 * 4. Builds per-attachment category names.
 * 5. Registers categories in Outlook (if not already present).
 * 6. Shows an info-bar notification.
 *
 * @param {Office.AddinCommands.Event} event  - The Office event object.
 * @param {string}                     trigger - Human-readable trigger reason (for logging).
 */
async function evaluateAndCategorize(event, trigger) {
  const item = Office.context.mailbox.item;

  // ── Step 1: Read To + CC ──────────────────────────────────────────────────
  const [toRecipients, ccRecipients] = await Promise.all([
    getRecipientsAsync(item.to),
    getRecipientsAsync(item.cc),
  ]);

  const allRecipients = [...toRecipients, ...ccRecipients];
  console.log(
    `[AttachCat] (${trigger}) Recipients:`,
    allRecipients.map((r) => r.emailAddress),
  );

  // ── Step 2: Find trigger matches ──────────────────────────────────────────
  const matchedLabels = getMatchedLabels(allRecipients);

  if (matchedLabels.length === 0) {
    // Remove any stale notification from a previous state
    item.notificationMessages.removeAsync(NOTIF_KEY_CATEGORIZED);
    item.notificationMessages.replaceAsync(NOTIF_KEY_NO_TRIGGER, {
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
      message: "No categorisation rules apply to the current recipients.",
      icon: "Icon.16x16",
      persistent: false,
    });
    return;
  }

  // ── Step 3: Read attachments ──────────────────────────────────────────────
  const attachments = await getAttachmentsAsync(item);

  if (attachments.length === 0) {
    item.notificationMessages.replaceAsync(NOTIF_KEY_CATEGORIZED, {
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
      message: `Trigger recipients detected (${matchedLabels.join(", ")}). Attach files to auto-categorize them.`,
      icon: "Icon.16x16",
      persistent: true,
    });
    return;
  }

  // ── Step 4: Build category names ─────────────────────────────────────────
  //  Category name format: "<RecipientLabel> | <DocType>"
  //  e.g. "Finance | Spreadsheet", "Legal | PDF Document"
  const categoryNames = buildCategoryNames(attachments, matchedLabels);
  console.log("[AttachCat] Derived categories:", categoryNames);

  // ── Step 5: Register & apply categories ──────────────────────────────────
  await ensureCategoriesExist(categoryNames, matchedLabels);
  await applyCategoriesAsync(item, categoryNames);

  // ── Step 6: Notify user ───────────────────────────────────────────────────
  const summary = categoryNames
    .map((cat, i) => `• ${attachments[i].name}  →  ${cat}`)
    .join("\n");

  item.notificationMessages.replaceAsync(NOTIF_KEY_CATEGORIZED, {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: `${attachments.length} attachment(s) categorized for ${matchedLabels.join(", ")}. Open taskpane for details.`,
    icon: "Icon.16x16",
    persistent: true,
  });

  // Store full detail in a custom property so the taskpane can read it
  await saveCategorizationSummaryAsync(
    item,
    attachments,
    categoryNames,
    matchedLabels,
  );
}

// ---------------------------------------------------------------------------
// ❹  HELPERS
// ---------------------------------------------------------------------------

/**
 * Wraps item.recipients.getAsync in a Promise.
 * Returns an empty array if the field is unavailable.
 */
function getRecipientsAsync(recipientField) {
  return new Promise((resolve) => {
    if (!recipientField || typeof recipientField.getAsync !== "function") {
      return resolve([]);
    }
    recipientField.getAsync({ asyncContext: null }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        console.warn("[AttachCat] Could not read recipients:", result.error);
        resolve([]);
      }
    });
  });
}

/**
 * Returns the array of category labels whose trigger email matches any
 * recipient in allRecipients.
 *
 * @param {Office.EmailAddressDetails[]} allRecipients
 * @returns {string[]}  deduplicated label array, e.g. ["Finance", "Legal"]
 */
function getMatchedLabels(allRecipients) {
  const matched = new Set();
  for (const recipient of allRecipients) {
    const email = (recipient.emailAddress || "").toLowerCase().trim();
    if (TRIGGER_RECIPIENT_MAP[email]) {
      matched.add(TRIGGER_RECIPIENT_MAP[email]);
    }
  }
  return [...matched];
}

/**
 * Wraps item.attachments (synchronous array) – returns a plain array or
 * falls back to getAttachmentsAsync for older API sets.
 */
function getAttachmentsAsync(item) {
  return new Promise((resolve) => {
    // attachments is a synchronous array in Mailbox 1.1+
    if (item.attachments) {
      return resolve(item.attachments);
    }
    // Fallback: use getAttachmentsAsync (Mailbox 1.8+)
    if (typeof item.getAttachmentsAsync === "function") {
      item.getAttachmentsAsync({}, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || []);
        } else {
          resolve([]);
        }
      });
    } else {
      resolve([]);
    }
  });
}

/**
 * Derives a category name for each attachment based on its file extension
 * and the matched recipient labels.
 *
 * If multiple labels matched the FIRST label is used to keep the name short.
 * You can change this to create one category per label × attachment if needed.
 *
 * @param {Office.AttachmentDetails[]} attachments
 * @param {string[]}                   matchedLabels
 * @returns {string[]}
 */
function buildCategoryNames(attachments, matchedLabels) {
  const primaryLabel = matchedLabels[0]; // first matched department
  return attachments.map((att) => {
    const ext = getExtension(att.name);
    const docType =
      ATTACHMENT_CATEGORY_RULES[ext] || ATTACHMENT_CATEGORY_RULES["_default"];
    // Combine ALL matched labels + doc type, e.g. "Finance & Legal | PDF Document"
    const labelStr =
      matchedLabels.length > 1 ? matchedLabels.join(" & ") : primaryLabel;
    return `${labelStr} | ${docType}`;
  });
}

/**
 * Extracts the lower-cased extension from a filename.
 * Returns "_default" if there is no extension.
 */
function getExtension(filename) {
  if (!filename) return "_default";
  const parts = filename.split(".");
  if (parts.length < 2) return "_default";
  return parts[parts.length - 1].toLowerCase();
}

/**
 * Registers any missing categories in the master category list so that
 * item.categories.addAsync can use them.
 *
 * Uses Office.context.mailbox.masterCategories (Mailbox 1.8).
 * Gracefully skips on older hosts that don't support it.
 */
async function ensureCategoriesExist(categoryNames, matchedLabels) {
  if (!Office.context.mailbox.masterCategories) return; // API not available

  const existing = await getMasterCategoriesAsync();
  const existingNames = new Set(existing.map((c) => c.displayName));

  const toCreate = [...new Set(categoryNames)]
    .filter((name) => !existingNames.has(name))
    .map((name) => {
      // Pick a colour based on which department label appears in the name
      const matchedLabel =
        matchedLabels.find((l) => name.startsWith(l)) || "_default";
      return {
        displayName: name,
        color:
          Office.MailboxEnums.CategoryColor[
            RECIPIENT_CATEGORY_COLORS[matchedLabel]
          ] || Office.MailboxEnums.CategoryColor.Preset8,
      };
    });

  if (toCreate.length === 0) return;

  return new Promise((resolve) => {
    Office.context.mailbox.masterCategories.addAsync(toCreate, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn(
          "[AttachCat] Could not register master categories:",
          result.error,
        );
      }
      resolve();
    });
  });
}

/** Wraps masterCategories.getAsync in a Promise. */
function getMasterCategoriesAsync() {
  return new Promise((resolve) => {
    Office.context.mailbox.masterCategories.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        resolve([]);
      }
    });
  });
}

/**
 * Applies the given category names to the mail item.
 * De-duplicates before adding so we don't accumulate stale duplicates
 * across multiple attachment-changed events.
 *
 * @param {Office.Item} item
 * @param {string[]}    categoryNames
 */
function applyCategoriesAsync(item, categoryNames) {
  // Deduplicate
  const unique = [...new Set(categoryNames)];
  return new Promise((resolve) => {
    item.categories.addAsync(unique, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn("[AttachCat] Could not apply categories:", result.error);
      }
      resolve();
    });
  });
}

/**
 * Saves a serialised summary to a custom property so the taskpane can
 * display the full per-attachment breakdown.
 *
 * @param {Office.Item}              item
 * @param {Office.AttachmentDetails[]} attachments
 * @param {string[]}                   categoryNames
 * @param {string[]}                   matchedLabels
 */
async function saveCategorizationSummaryAsync(
  item,
  attachments,
  categoryNames,
  matchedLabels,
) {
  const summary = {
    timestamp: new Date().toISOString(),
    matchedLabels,
    attachments: attachments.map((att, i) => ({
      id: att.id,
      name: att.name,
      size: att.size,
      category: categoryNames[i],
    })),
  };

  return new Promise((resolve) => {
    item.loadCustomPropertiesAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded)
        return resolve();
      const props = result.value;
      props.set("attachCatSummary", JSON.stringify(summary));
      props.saveAsync(() => resolve());
    });
  });
}

// ---------------------------------------------------------------------------
// ❺  REGISTER HANDLERS WITH OFFICE RUNTIME
// ---------------------------------------------------------------------------
//  Required so Office can resolve the function names declared in the manifest.

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
}
