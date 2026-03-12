# Outlook Attachment Categorizer Add-in

Auto-categorizes attachments in a compose window when specific recipients
are detected in the **To** or **CC** fields.

---

## How it works

| Step | What happens |
|------|--------------|
| 1 | User opens a new compose window — add-in loads silently via event-based activation |
| 2 | User adds a recipient that matches a trigger address (e.g. `finance@contoso.com`) |
| 3 | **OnMessageRecipientsChanged** fires → add-in notes the trigger |
| 4 | User attaches a file |
| 5 | **OnMessageAttachmentsChanged** fires → add-in reads every attachment, derives a category name (`Finance \| PDF Document`) and applies it to the item |
| 6 | An info-bar notification confirms which categories were applied |
| 7 | User can click **View Categories** in the ribbon to open the taskpane for a full breakdown |

---

## File layout

```
outlook-attachment-categorizer/
├── manifest.xml                          ← Sideload this
├── package.json
└── src/
    ├── commands/
    │   ├── commands.js                   ← All event handlers + core logic
    │   └── commands.html                 ← WebView runtime host (silent)
    └── taskpane/
        └── taskpane.html                 ← Visual breakdown (ribbon button)
```

---

## Configuration

Edit the two objects at the top of `src/commands/commands.js`
(and the identical copies in `taskpane.html`):

### 1 — Trigger recipients

```js
const TRIGGER_RECIPIENT_MAP = {
  "finance@contoso.com":    "Finance",
  "legal@contoso.com":      "Legal",
  "hr@contoso.com":         "HR",
  "compliance@contoso.com": "Compliance",
  "audit@contoso.com":      "Audit",
};
```

Add or remove entries as needed.
The value (e.g. `"Finance"`) becomes the prefix of every category name.

### 2 — Attachment → document-type rules

```js
const ATTACHMENT_CATEGORY_RULES = {
  pdf:  "PDF Document",
  docx: "Word Document",
  xlsx: "Spreadsheet",
  // …extend as needed
  _default: "General Attachment",
};
```

### Category name format

`<RecipientLabel(s)> | <DocType>`

Examples:
- `Finance | PDF Document`
- `Legal & HR | Spreadsheet`
- `Compliance | General Attachment`

---

## Setup & sideloading

### Prerequisites
- Node.js ≥ 16
- Microsoft 365 account with Outlook (web, desktop, or Mac)

### Run locally

```bash
npm install
npm start          # installs a trusted dev SSL cert + starts https://localhost:3000
```

### Sideload the manifest

**Outlook on the Web**
1. Open Outlook on the web → Settings → Integrated Apps → Upload custom app
2. Select `manifest.xml`

**Classic Outlook on Windows**
1. File → Manage Add-ins → My Add-ins → Add a custom add-in → Add from file
2. Select `manifest.xml`

**Outlook on Mac**
1. Toolbar → More → Get Add-ins → My Add-ins → + → Add from file
2. Select `manifest.xml`

---

## Deployment

Replace all `https://localhost:3000` URLs in `manifest.xml` with your production
hosting URL, then deploy the `src/` folder to that host (Azure Static Web Apps,
GitHub Pages, etc.).

---

## Mailbox API requirements

| Feature | Minimum API set |
|---------|----------------|
| Event-based activation | Mailbox 1.10 |
| `item.categories.addAsync` | Mailbox 1.8 |
| `masterCategories.addAsync` | Mailbox 1.8 |
| `item.attachments` (sync) | Mailbox 1.1 |

The manifest declares `Mailbox 1.3` as the minimum for VersionOverrides so the
add-in can be installed broadly; the code feature-detects and gracefully skips
APIs that aren't available on older hosts.
