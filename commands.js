// // // =============================================================================
// // //  commands.js  – Attachment Categorizer (Prefix-Rename Flow)
// // //
// // //  Events:
// // //    OnNewMessageCompose            → check pre-filled recipients
// // //    OnMessageAttachmentsChanged    → prompt user to categorize when trigger recipient present
// // //    OnMessageRecipientsChanged     → re-evaluate when To/CC changes
// // // =============================================================================
// // "use strict";

// // // ── CONFIGURATION ─────────────────────────────────────────────────────────────

// // const TRIGGER_RECIPIENT_MAP = {
// //   "abhishek.a3@99acres.com": "Abhishek Anand",
// //   "sonia.m@99acres.com": "Sonia M",
// //   "coder.abhi02@gmail.com": "Abhishek Kumar",
// //   "rkd02122@gmail.com": "Abhishek Kumar",
// //   "finance@contoso.com": "Finance",
// //   "legal@contoso.com": "Legal",
// //   "hr@contoso.com": "HR",
// //   "compliance@contoso.com": "Compliance",
// //   "audit@contoso.com": "Audit",
// // };

// // const NOTIF_KEY_ACTION = "attachCatAction";
// // const NOTIF_KEY_NO_RULE = "attachCatNoRule";
// // const TASKPANE_URL_RESID = "Taskpane.Url";

// // // ── ENTRY POINTS ──────────────────────────────────────────────────────────────

// // async function onNewMessageComposeHandler(event) {
// //   try {
// //     await checkAndNotify();
// //   } catch (e) {
// //     console.error("[AttachCat]", e);
// //   } finally {
// //     event.completed();
// //   }
// // }

// // async function onMessageAttachmentsChangedHandler(event) {
// //   try {
// //     await checkAndNotify();
// //   } catch (e) {
// //     console.error("[AttachCat]", e);
// //   } finally {
// //     event.completed();
// //   }
// // }

// // async function onMessageRecipientsChangedHandler(event) {
// //   try {
// //     await checkAndNotify();
// //   } catch (e) {
// //     console.error("[AttachCat]", e);
// //   } finally {
// //     event.completed();
// //   }
// // }

// // // ── SEND GUARD ─────────────────────────────────────────────────────────────────
// // /**
// //  * OnMessageSend fires when the user clicks Send.
// //  * - If no trigger recipients → allow send (not our concern).
// //  * - If trigger recipients present but attachments exist without a category prefix → BLOCK.
// //  * - If all attachments are categorized (or no attachments) → allow send.
// //  *
// //  * SendMode="SoftBlock" in manifest means the user sees:
// //  *   "Don't send" | "Send Anyway"
// //  * Change to SendMode="Block" in manifest for a hard block with no bypass.
// //  */
// // async function onMessageSendHandler(event) {
// //   try {
// //     const item = Office.context.mailbox.item;

// //     // 1. Check for trigger recipients
// //     const [toR, ccR] = await Promise.all([
// //       getRecipientsAsync(item.to),
// //       getRecipientsAsync(item.cc),
// //     ]);
// //     const matchedLabels = getMatchedLabels([...toR, ...ccR]);

// //     // No trigger recipients → not our email, let it through
// //     if (matchedLabels.length === 0) {
// //       event.completed({ allowEvent: true });
// //       return;
// //     }

// //     // 2. Check attachments
// //     const attachments = item.attachments || [];

// //     // No attachments at all → warn but allow (they may be sending text-only)
// //     if (attachments.length === 0) {
// //       event.completed({ allowEvent: true });
// //       return;
// //     }

// //     // 3. Find uncategorized attachments
// //     //    Categorized = filename starts with a known category prefix e.g. "Legal_"
// //     const KNOWN_PREFIXES = [
// //       "Legal_",
// //       "Finance_",
// //       "HR_",
// //       "Compliance_",
// //       "Contract_",
// //       "Invoice_",
// //       "Report_",
// //       "Presentation_",
// //       "Reference_",
// //       "General_",
// //     ];

// //     const uncategorized = attachments.filter(
// //       (att) => !KNOWN_PREFIXES.some((prefix) => att.name.startsWith(prefix)),
// //     );

// //     if (uncategorized.length > 0) {
// //       const names = uncategorized.map((a) => `• ${a.name}`).join("\n");

// //       // Block the send — Office shows a dialog with the errorMessage
// //       // SoftBlock:  user sees "Don't Send" | "Send Anyway"
// //       // Block:      user sees only "Don't Send"
// //       event.completed({
// //         allowEvent: false,
// //         cancelLabel: "Go Back & Categorize", // label for the "don't send" button
// //         commandId: "msgComposeOpenPaneButton", // opens taskpane when user clicks it
// //         contextData: JSON.stringify({
// //           uncategorized: uncategorized.map((a) => a.name),
// //         }),
// //         errorMessage:
// //           `${uncategorized.length} attachment(s) are not categorized:\n\n${names}\n\n` +
// //           `Please open "View Categories" in the ribbon and label each file before sending.`,
// //       });
// //       return;
// //     }

// //     // All categorized → allow
// //     event.completed({ allowEvent: true });
// //   } catch (e) {
// //     console.error("[AttachCat] onMessageSend error:", e);
// //     // On unexpected error, allow send so we don't permanently block the user
// //     event.completed({ allowEvent: true });
// //   }
// // }

// // // ── CORE ──────────────────────────────────────────────────────────────────────

// // async function checkAndNotify() {
// //   const item = Office.context.mailbox.item;

// //   const [toR, ccR] = await Promise.all([
// //     getRecipientsAsync(item.to),
// //     getRecipientsAsync(item.cc),
// //   ]);

// //   const allRecipients = [...toR, ...ccR];
// //   const matchedLabels = getMatchedLabels(allRecipients);

// //   // No trigger recipient → clear and exit
// //   if (matchedLabels.length === 0) {
// //     item.notificationMessages.removeAsync(NOTIF_KEY_ACTION);
// //     item.notificationMessages.replaceAsync(NOTIF_KEY_NO_RULE, {
// //       type: Office.MailboxEnums.ItemNotificationMessageType
// //         .InformationalMessage,
// //       message: "No categorisation rules apply to current recipients.",
// //       icon: "Icon.16x16",
// //       persistent: false,
// //     });
// //     return;
// //   }

// //   // Check attachments
// //   const attachments = item.attachments || [];

// //   if (attachments.length === 0) {
// //     item.notificationMessages.removeAsync(NOTIF_KEY_NO_RULE);
// //     item.notificationMessages.replaceAsync(NOTIF_KEY_ACTION, {
// //       type: Office.MailboxEnums.ItemNotificationMessageType
// //         .InformationalMessage,
// //       message: `Sending to ${matchedLabels.join(", ")} — attach files and open the categorizer to label them.`,
// //       icon: "Icon.16x16",
// //       persistent: true,
// //     });
// //     return;
// //   }

// //   // Attachments present → prompt to categorize
// //   const uncategorized = attachments.filter((a) => !a.name.includes("_"));
// //   const msg =
// //     uncategorized.length > 0
// //       ? `${uncategorized.length} attachment(s) need categorization. Open "View Categories" to label them.`
// //       : `All ${attachments.length} attachment(s) categorized. Open "View Categories" to review.`;

// //   item.notificationMessages.removeAsync(NOTIF_KEY_NO_RULE);
// //   item.notificationMessages.replaceAsync(NOTIF_KEY_ACTION, {
// //     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
// //     message: msg,
// //     icon: "Icon.16x16",
// //     persistent: true,
// //   });
// // }

// // // ── HELPERS ───────────────────────────────────────────────────────────────────

// // function getRecipientsAsync(field) {
// //   return new Promise((resolve) => {
// //     if (!field || typeof field.getAsync !== "function") return resolve([]);
// //     field.getAsync({}, (result) =>
// //       resolve(
// //         result.status === Office.AsyncResultStatus.Succeeded
// //           ? result.value || []
// //           : [],
// //       ),
// //     );
// //   });
// // }

// // function getMatchedLabels(recipients) {
// //   const matched = new Set();
// //   for (const r of recipients) {
// //     const email = (r.emailAddress || "").toLowerCase().trim();
// //     if (TRIGGER_RECIPIENT_MAP[email]) matched.add(TRIGGER_RECIPIENT_MAP[email]);
// //   }
// //   return [...matched];
// // }

// // // ── REGISTER ──────────────────────────────────────────────────────────────────
// // if (typeof Office !== "undefined") {
// //   Office.actions.associate(
// //     "onNewMessageComposeHandler",
// //     onNewMessageComposeHandler,
// //   );
// //   Office.actions.associate(
// //     "onMessageAttachmentsChangedHandler",
// //     onMessageAttachmentsChangedHandler,
// //   );
// //   Office.actions.associate(
// //     "onMessageRecipientsChangedHandler",
// //     onMessageRecipientsChangedHandler,
// //   );
// //   Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
// // }
// // =============================================================================
// //  commands.js  – Attachment Categorizer (Prefix-Rename Flow)
// //
// //  Events:
// //    OnNewMessageCompose            → check pre-filled recipients
// //    OnMessageAttachmentsChanged    → prompt user to categorize when trigger recipient present
// //    OnMessageRecipientsChanged     → re-evaluate when To/CC changes
// // =============================================================================
// "use strict";

// // ── CONFIGURATION ─────────────────────────────────────────────────────────────

// const TRIGGER_RECIPIENT_MAP = {
//   "abhishek.a3@99acres.com": "Abhishek Anand",
//   "sonia.m@99acres.com": "Sonia M",
//   "coder.abhi02@gmail.com": "Abhishek Kumar",
//   "rkd02122@gmail.com": "Abhishek Kumar",
//   "finance@contoso.com": "Finance",
//   "legal@contoso.com": "Legal",
//   "hr@contoso.com": "HR",
//   "compliance@contoso.com": "Compliance",
//   "audit@contoso.com": "Audit",
// };

// const NOTIF_KEY_ACTION = "attachCatAction";
// const NOTIF_KEY_NO_RULE = "attachCatNoRule";
// const TASKPANE_URL_RESID = "Taskpane.Url";

// // ── ENTRY POINTS ──────────────────────────────────────────────────────────────

// async function onNewMessageComposeHandler(event) {
//   try {
//     await checkAndNotify();
//   } catch (e) {
//     console.error("[AttachCat]", e);
//   } finally {
//     event.completed();
//   }
// }

// async function onMessageAttachmentsChangedHandler(event) {
//   try {
//     await checkAndNotify();
//   } catch (e) {
//     console.error("[AttachCat]", e);
//   } finally {
//     event.completed();
//   }
// }

// async function onMessageRecipientsChangedHandler(event) {
//   try {
//     await checkAndNotify();
//   } catch (e) {
//     console.error("[AttachCat]", e);
//   } finally {
//     event.completed();
//   }
// }

// // ── SEND GUARD ────────────────────────────────────────────────────────────────
// /**
//  * OnMessageSend fires when the user clicks Send.
//  * MUST complete within 5 seconds — use only synchronous item properties.
//  *
//  * SendMode="SoftBlock" → user sees "Don't Send" | "Send Anyway"
//  * SendMode="Block"     → hard block, no bypass
//  */

// // function onMessageSendHandler(event) {
// //   try {
// //     const item = Office.context.mailbox.item;

// //     // ── Use synchronous item.attachments (available at send time) ────────────
// //     const attachments = item.attachments || [];

// //     // No attachments → nothing to categorize, allow send
// //     if (attachments.length === 0) {
// //       event.completed({ allowEvent: true });
// //       return;
// //     }

// //     // ── Check each attachment for a known category prefix ────────────────────
// //     const KNOWN_PREFIXES = [
// //       "Legal_",
// //       "Finance_",
// //       "HR_",
// //       "Compliance_",
// //       "Contract_",
// //       "Invoice_",
// //       "Report_",
// //       "Presentation_",
// //       "Reference_",
// //       "General_",
// //     ];

// //     const uncategorized = attachments.filter(
// //       (att) => !KNOWN_PREFIXES.some((p) => att.name.startsWith(p)),
// //     );

// //     if (uncategorized.length === 0) {
// //       // All categorized → allow send
// //       event.completed({ allowEvent: true });
// //       return;
// //     }

// //     // ── Block send with clear message ─────────────────────────────────────────
// //     const fileList = uncategorized.map((a) => `• ${a.name}`).join("\n");

// //     event.completed({
// //       allowEvent: false,
// //       errorMessage:
// //         `${uncategorized.length} attachment(s) are not categorized:\n\n` +
// //         `${fileList}\n\n` +
// //         `Please open "View Categories" in the ribbon and label each file before sending.`,
// //     });
// //   } catch (e) {
// //     console.error("[AttachCat] onMessageSend error:", e);
// //     // On unexpected error, allow send so user is never permanently stuck
// //     event.completed({ allowEvent: true });
// //   }
// // }

// // ── CORE ──────────────────────────────────────────────────────────────────────

// function onMessageSendHandler(event) {
//   const attachments = Office.context.mailbox.item.attachments || [];

//   // No attachments → nothing to check
//   if (attachments.length === 0) {
//     event.completed({ allowEvent: true });
//     return;
//   }

//   const uncategorized = attachments.filter(
//     (att) => !KNOWN_PREFIXES.some((p) => att.name.startsWith(p)),
//   );

//   if (uncategorized.length === 0) {
//     event.completed({ allowEvent: true });
//     return;
//   }

//   // Block send immediately
//   event.completed({
//     allowEvent: false,
//     errorMessage:
//       `⚠️ ${uncategorized.length} attachment(s) not yet categorized:\n\n` +
//       uncategorized.map((a) => `• ${a.name}`).join("\n") +
//       `\n\nPlease open "View Categories" in the ribbon and label each file before sending.`,
//   });
// }

// async function checkAndNotify() {
//   const item = Office.context.mailbox.item;

//   const [toR, ccR] = await Promise.all([
//     getRecipientsAsync(item.to),
//     getRecipientsAsync(item.cc),
//   ]);

//   const allRecipients = [...toR, ...ccR];
//   const matchedLabels = getMatchedLabels(allRecipients);

//   // No trigger recipient → clear and exit
//   if (matchedLabels.length === 0) {
//     item.notificationMessages.removeAsync(NOTIF_KEY_ACTION);
//     item.notificationMessages.replaceAsync(NOTIF_KEY_NO_RULE, {
//       type: Office.MailboxEnums.ItemNotificationMessageType
//         .InformationalMessage,
//       message: "No categorisation rules apply to current recipients.",
//       icon: "Icon.16x16",
//       persistent: false,
//     });
//     return;
//   }

//   // Check attachments
//   const attachments = item.attachments || [];

//   if (attachments.length === 0) {
//     item.notificationMessages.removeAsync(NOTIF_KEY_NO_RULE);
//     item.notificationMessages.replaceAsync(NOTIF_KEY_ACTION, {
//       type: Office.MailboxEnums.ItemNotificationMessageType
//         .InformationalMessage,
//       message: `Sending to ${matchedLabels.join(", ")} — attach files and open the categorizer to label them.`,
//       icon: "Icon.16x16",
//       persistent: true,
//     });
//     return;
//   }

//   // Attachments present → prompt to categorize
//   const uncategorized = attachments.filter((a) => !a.name.includes("_"));
//   const msg =
//     uncategorized.length > 0
//       ? `${uncategorized.length} attachment(s) need categorization. Open "View Categories" to label them.`
//       : `All ${attachments.length} attachment(s) categorized. Open "View Categories" to review.`;

//   item.notificationMessages.removeAsync(NOTIF_KEY_NO_RULE);
//   item.notificationMessages.replaceAsync(NOTIF_KEY_ACTION, {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: msg,
//     icon: "Icon.16x16",
//     persistent: true,
//   });
// }

// // ── HELPERS ───────────────────────────────────────────────────────────────────

// function getRecipientsAsync(field) {
//   return new Promise((resolve) => {
//     if (!field || typeof field.getAsync !== "function") return resolve([]);
//     field.getAsync({}, (result) =>
//       resolve(
//         result.status === Office.AsyncResultStatus.Succeeded
//           ? result.value || []
//           : [],
//       ),
//     );
//   });
// }

// function getMatchedLabels(recipients) {
//   const matched = new Set();
//   for (const r of recipients) {
//     const email = (r.emailAddress || "").toLowerCase().trim();
//     if (TRIGGER_RECIPIENT_MAP[email]) matched.add(TRIGGER_RECIPIENT_MAP[email]);
//   }
//   return [...matched];
// }

// // ── REGISTER ──────────────────────────────────────────────────────────────────
// if (typeof Office !== "undefined") {
//   Office.actions.associate(
//     "onNewMessageComposeHandler",
//     onNewMessageComposeHandler,
//   );
//   Office.actions.associate(
//     "onMessageAttachmentsChangedHandler",
//     onMessageAttachmentsChangedHandler,
//   );
//   Office.actions.associate(
//     "onMessageRecipientsChangedHandler",
//     onMessageRecipientsChangedHandler,
//   );
//   Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
// }

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
//  SEND GUARD — synchronous, no await, no API calls
//  Office gives OnMessageSend only 5 seconds — this runs in <1ms
// =============================================================================
function onMessageSendHandler(event) {
  const attachments = Office.context.mailbox.item.attachments || [];

  if (attachments.length === 0) {
    event.completed({ allowEvent: true });
    return;
  }

  const uncategorized = attachments.filter(
    (att) => !KNOWN_PREFIXES.some((p) => att.name.startsWith(p)),
  );

  if (uncategorized.length === 0) {
    event.completed({ allowEvent: true });
    return;
  }

  event.completed({
    allowEvent: false,
    errorMessage:
      "⚠️ " +
      uncategorized.length +
      " attachment(s) not categorized:\n\n" +
      uncategorized.map((a) => "• " + a.name).join("\n") +
      "\n\nOpen 'View Categories' in the ribbon to label each file before sending.",
  });
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

  const attachments = item.attachments || [];
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
