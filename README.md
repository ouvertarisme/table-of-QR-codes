# Google Docs — Footnote URL QR Code Table

Turn footnote links in a Google Doc into a clean, printable **QR Code Index**.

This Apps Script:
- Scans **all footnotes** and extracts URLs (linked text or plain-text).
- Fetches each page’s HTML `<title>` (best effort, with safe fallback).
- Generates a **3‑column table** at your chosen placeholder location (**`QRCodeTable`**):
  - **Ref** — 2‑digit hex code `01..FF` (sequential).
  - **Title** (line 1, size 11) and **URL** (line 2, size 8, single line spanning under Title+QR).
  - **QR Code** image that points to the URL (vertically centered next to the title).
- Layout per entry uses **two sub‑rows**:
  - Top sub‑row: Ref | Title | QR
  - Bottom sub‑row: (blank) | URL **merged across 2 columns**

> **Why this repo?** Handy for books, white papers, reports—where readers can scan long references from print/PDF.

---

## ✨ Features
- **Deterministic refs**: `01`, `02`, … (up to `FF`), i.e. no more than 255 QR codes! Duplicate URLs are deduped by default.
- **Robust QR generation** with fallbacks: QuickChart → goQR → ZXing.
- **High‑quality print**: generate higher‑resolution QR images and display smaller in Doc.
- **Marker‑based injection**: table is inserted **in place of** the string `QRCodeTable`. If the marker is missing, the script **does nothing**.

---

## 🧩 How it works (high level)
1. The script walks the Doc body to find all **Footnotes** and extracts URLs.
2. For each URL, it `UrlFetchApp.fetch()`‑es the target and parses the `<title>` if present.
3. It searches the Doc for the literal text **`QRCodeTable`**. If not found → aborts.
4. It inserts a 3‑column table, then uses the **Advanced Google Docs API** to:
   - Vertically merge the **Ref** cells across the two sub‑rows,
   - Horizontally merge the **URL** across the two columns in the bottom sub‑row,
   - Vertically center **Ref**, **Title**, and **QR** cells.

---

## 🚀 Quick Start

### 1) Copy code into your Google Doc
1. Open your target **Google Doc**.
2. Go to **Extensions → Apps Script**.
3. In `Code.gs`, paste the script from `src/Code.gs` (or your copy).
4. Save.

### 2) Enable the Advanced Google Docs API
Google Apps Script’s built‑in `DocumentApp` can’t merge cells or set vertical alignment, so we use the Advanced Docs API.

- In Apps Script editor, click the **puzzle icon (Services)** → **+** → add **Google Docs API**.
- (Optional) In **Project Settings → Google Cloud Platform (GCP) Project**, open the Console and ensure **Google Docs API** is enabled. Usually adding the service is enough.

### 3) Place the placeholder in your Doc
Insert the exact marker text where you want the table to appear:

```
QRCodeTable
```

> The script **replaces** this exact text with the QR Code table. If not present, it will **abort** without modifying the Doc.

### 4) Run it
- Back in the Doc, a custom menu **QR Tools** appears (if not, reload the Doc).
- Choose **QR Tools → Build QR Table**.
- First run: accept the authorization prompts.

That’s it!

---

## ⚙️ Configuration

You can adjust the following in the script:

- **QR base size** (requested from endpoints):
  ```js
  const blob = buildQrBlob_(e.url, 600); // 600 px recommended for crisp print
  ```
- **On‑page size** (displayed smaller for higher effective PPI):
  ```js
  const img = qrCell.insertImage(0, blob);
  img.setWidth(Math.floor(img.getWidth() * 0.33));
  img.setHeight(Math.floor(img.getHeight() * 0.33));
  ```
  > Display scaling changes visual size **without** changing the image’s pixel count. For print/PDF, you get higher effective PPI.
- **De‑duplication**: remove or change the `Set` in `extractUrlsFromAllFootnotes_` to list duplicate links more than once.
- **Retry pacing** for QR endpoints:
  ```js
  Utilities.sleep(60); // ms between endpoint attempts
  ```
- **Title fallback**:
  ```js
  const title = fetchPageTitle_(u) || 'Title not available — please edit';
  ```

---

## 🔒 Permissions & Scopes

When you authorize the script, it will request access to:
- Read/edit the current Doc (to read footnotes and insert the table).
- External service access for `UrlFetchApp` (to fetch titles and QR images).
- Advanced Google Docs API (to merge cells and set vertical alignment).

No data is stored by the script; HTTP calls are made only to the link targets and QR endpoints you configure.

---

## 🧪 Troubleshooting

**No menu appears**
- Reload the Doc to trigger `onOpen()`.

**“Placeholder not found”**
- Ensure the Doc contains the exact text `QRCodeTable` (case‑sensitive).

**Titles are empty or odd characters**
- Some sites restrict fetching or use unusual encodings. The script falls back to a placeholder; edit it manually in the table if needed.

**QR codes fail to generate**
- The script tries QuickChart → goQR → ZXing with short delays for rate limits.
- If you’re generating **hundreds** at once, consider slightly increasing the sleep or batching your runs.
- Firewalls or network policies can block external endpoints; pick one that’s reachable from your environment.

**URLs wrap to multiple lines**
- The bottom row is merged across two columns and uses 8pt font to maximize space. Very long URLs can still wrap due to Docs’ layout rules. Consider using short links if print layout is critical.

---

## 📦 Files

- `src/Code.gs` — main Apps Script file (menu, scanning, table building, QR generator, merges/styles).
- `README.md` — this file.

> You don’t need a `manifest.json`; Apps Script manages it for simple projects. The Advanced Service toggle stores the dependency.

---

## 🤝 Contributing

PRs welcome! Please keep the code:
- Vanilla Apps Script (no third‑party libs).
- Readable and well‑commented.
- Safe on quotas (avoid aggressive parallel fetching).

---

## 📝 License

This project is released under **The Unlicense** (public domain). See `LICENSE` or <https://unlicense.org/>.

---

## 📬 Support

Open an issue with:
- A redacted test Doc (structure + a couple of footnotes),
- The Apps Script execution log (errors),
- The exact behavior you observed.

Happy scanning! 📎➡️📱
