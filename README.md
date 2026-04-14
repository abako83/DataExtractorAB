# DataSnipper AI — Excel Add-in

An Excel task pane add-in that uses Claude's vision API to extract data from PDFs and images directly into your spreadsheet.

## Features

- Upload **PDFs or images** (PNG, JPG, WEBP) via click or drag-and-drop
- Extract as **Table**, **Key-Value pairs**, **List**, or **Raw text**
- Custom instructions (e.g. "only extract invoice line items")
- Auto-inserts with **green header formatting** at your selected cell
- API key stored locally in the browser — never sent anywhere except Anthropic

---

## Quick Setup (Local Dev)

### 1. Prerequisites
- Node.js 18+
- Excel (Microsoft 365 desktop or Excel on the web)
- An [Anthropic API key](https://console.anthropic.com)

### 2. Install a local HTTPS server

```bash
npm install -g office-addin-dev-certs http-server
```

Generate a local SSL cert (required by Office):
```bash
office-addin-dev-certs install
```

### 3. Serve the add-in files

```bash
cd datasnipper-addin
http-server . -p 3000 --ssl --cert ~/.office-addin-dev-certs/localhost.crt --key ~/.office-addin-dev-certs/localhost.key
```

Your add-in is now at `https://localhost:3000/taskpane.html`

### 4. Update the manifest

In `manifest.xml`, replace all instances of `https://your-server.com` with `https://localhost:3000`

### 5. Sideload into Excel

**Excel on the web (easiest):**
1. Open any workbook on [office.com](https://office.com)
2. Insert → Add-ins → Upload My Add-in
3. Browse to `manifest.xml` and upload

**Excel Desktop (Windows):**
1. Copy `manifest.xml` to a shared network folder
2. Excel → File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs
3. Add the folder path → OK
4. Insert → My Add-ins → Shared Folder → DataSnipper AI

**Excel Desktop (Mac):**
1. Copy `manifest.xml` to `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
2. Restart Excel → Insert → My Add-ins

---

## Production Deployment

1. Deploy `taskpane.html` to any HTTPS host (Vercel, Netlify, Azure Static Web Apps, S3+CloudFront)
2. Update all `https://your-server.com` URLs in `manifest.xml`
3. Submit to [Microsoft AppSource](https://partner.microsoft.com/dashboard) or distribute the manifest directly to your org via Microsoft 365 Admin Center

---

## Architecture

```
Excel (Office JS)
    ↓  task pane loads
taskpane.html  (vanilla JS + Office JS SDK)
    ↓  sends base64 PDF/image
Anthropic API  /v1/messages  (claude-sonnet-4-20250514)
    ↓  returns JSON or text
Excel worksheet  (inserted at selected cell)
```

## Extraction Modes

| Mode | Output | JSON shape |
|------|--------|------------|
| Table | Tabular data with headers | `{headers, rows}` |
| Key-Value | Label + value pairs | `{pairs: [{key, value}]}` |
| List | Enumerated items | `{items: [...]}` |
| Raw text | Plain transcription | string |

## Files

| File | Purpose |
|------|---------|
| `taskpane.html` | The full add-in UI and logic (single file) |
| `manifest.xml` | Office add-in manifest — registers the add-in with Excel |
| `README.md` | This file |

---

## Customization Tips

- **Change the model**: Edit `claude-sonnet-4-20250514` in `taskpane.html` to use `claude-opus-4-20250514` for more complex docs
- **Add more modes**: Duplicate a mode chip and add a matching prompt in `modeInstructions`
- **Auto-select a range**: After insertion, the code formats headers green — adjust colors in `insertToSheet()`
- **Save API key to server**: Replace `localStorage` with a call to your own auth backend
