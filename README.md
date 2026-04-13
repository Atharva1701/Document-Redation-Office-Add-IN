# Document Redaction Add-in for Microsoft Word

A Microsoft Word task-pane add-in that automatically redacts sensitive information from documents, inserts a **CONFIDENTIAL DOCUMENT** header, and optionally enables Track Changes — all in one click.

---

## Features

- **Sensitive data redaction** across all document paragraphs:
  - Email addresses → `[EMAIL REDACTED]`
  - Phone numbers (US formats) → `[PHONE REDACTED]`
  - Social Security Numbers → `[SSN REDACTED]`
  - Credit card numbers (Luhn-validated, 13–19 digits) → `[CREDIT CARD REDACTED]`
- **Confidential header** — inserts `CONFIDENTIAL DOCUMENT` at the top of the document (idempotent; skipped if already present)
- **Track Changes** — enables revision tracking via the Word JavaScript API when the host supports it (WordApi 1.5 + WordApiDesktop 1.4)
- **Redaction summary** — displays a per-category count of all redactions made and reports the tracking/header status

---

## Tech Stack

| Layer | Technology |
|---|---|
| Language | TypeScript 5.6 |
| Build tool | Vite 5 |
| Office integration | Office.js (Word JavaScript API) |
| Styling | Vanilla CSS (custom properties) |
| HTTPS (dev) | `@vitejs/plugin-basic-ssl` / local cert |

---

## Project Structure

```
├── manifest.xml                   
├── public/
│   ├── icon-32.png
│   └── icon-64.png
├── src/
│   ├── redactor/
│   │   ├── luhn.ts                # Luhn algorithm for credit-card validation
│   │   ├── patterns.ts            # Compiled regex patterns (email, phone, SSN, card)
│   │   └── redactor.ts            # Core redaction engine (Word.run orchestrator)
│   ├── taskpane/
│   │   ├── taskpane.html          # Add-in task-pane UI
│   │   ├── taskpane.ts            # Task-pane controller (Office.onReady, event wiring)
│   │   └── taskpane.css           # Task-pane styles
├── tsconfig.json
├── vite.config.ts
└── package.json
```

---

## Prerequisites

- **Node.js** ≥ 18
- **npm** ≥ 9
- **Microsoft Word** — desktop (Windows/Mac) or Word on the Web
- A local HTTPS certificate pair (`localhost.pem` + `localhost-key.pem`) in the project root

### Generating a local certificate

If you do not already have a trusted local certificate, use [mkcert](https://github.com/FiloSottile/mkcert):

```bash
# Install mkcert (macOS example)
brew install mkcert
mkcert -install

# Generate certs in the project root
cd <your-project-directory>
mkcert localhost
```

This produces `localhost.pem` and `localhost-key.pem`, which are referenced by `vite.config.ts`.

---

## Getting Started

```bash
# 1. Clone the repository
git clone <your-repo-url>
cd <your-project-name>

# 2. Install dependencies
npm install

# 3. Start the development server (HTTPS, port 3000)
npm run dev
```

Then sideload the add-in into Word using `manifest.xml` (see [Sideloading](#sideloading) below).

### Available scripts

| Script | Description |
|---|---|
| `npm run dev` | Start Vite dev server on `https://localhost:3000` |
| `npm run build` | Production build to `dist/` |
| `npm run preview` | Serve the production build locally |
| `npm run typecheck` | Type-check without emitting files |

---

## Sideloading

If automatic sideloading does not trigger, follow Microsoft's manual sideload guide:

- **Word on the Web** — [Sideload Office Add-ins in Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#sideload-an-office-add-in-in-office-on-the-web)
- **Word Desktop (Windows)** — [Sideload Office Add-ins on Windows](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#sideload-an-office-add-in-on-windows)
- **Word Desktop (Mac)** — [Sideload Office Add-ins on Mac](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#sideload-an-office-add-in-on-mac)

---

## Usage

1. Open any `.docx` file in Microsoft Word.
2. Open the **Document Redaction** task pane from the add-in.
3. Click **Redact document**.
4. A summary card appears showing how many items of each type were redacted.

The add-in is idempotent — running it multiple times on an already-redacted document will not produce duplicate headers or double-redact tokens.

---

## How It Works

### Redaction pipeline (`src/redactor/redactor.ts`)

1. **Track Changes** — attempts to enable `document.trackRevisions` if WordApi 1.5 and WordApiDesktop 1.4 requirement sets are both available on the current host.
2. **Confidential header** — reads the document body text and inserts a `CONFIDENTIAL DOCUMENT` paragraph at the start only if it is not already present.
3. **Pattern matching** — iterates every paragraph, runs four regex passes (email → phone → SSN → card candidate), and replaces matches in-place via `paragraph.insertText(…, Word.InsertLocation.replace)`.
4. **Credit card validation** — card-shaped digit sequences are passed through a Luhn checksum (`src/redactor/luhn.ts`) before being redacted, minimising false positives.

### Regex patterns (`src/redactor/patterns.ts`)

| Type | Pattern highlights |
|---|---|
| Email | RFC-5321 local part + domain, case-insensitive |
| Phone | US 10-digit numbers with optional country code, various separators |
| SSN | Strict `\d{3}-\d{2}-\d{4}` format |
| Credit card | 13–19 digit sequences (spaces/dashes allowed), Luhn-validated |

---

## Word API Requirements

| Feature | Requirement set |
|---|---|
| Task pane, document read/write | WordApi 1.1 |
| `trackRevisions` | WordApi 1.5 + WordApiDesktop 1.4 |

Track Changes is gracefully degraded — if the host does not support the required sets, redaction and header insertion proceed normally and a note is shown in the summary.

Reference: [Word JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

## Testing

You can test this add-in on your documents (.docx) files.

> **Note:** The redaction logic is document-agnostic. Ensure your patterns cover the formats present in any document you intend to process.

---

