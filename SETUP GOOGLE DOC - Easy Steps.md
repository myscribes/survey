# SETUP — Kristy Google Doc (Responses Master)

**Kristy only — not connected to Melanie.**

When someone submits the survey, each response is saved to:
1. **Google Sheet** (rows)
2. **Google Doc** (formatted, newest first) ← you are setting this up now

---

## PART 1 — Create the Google Doc

1. Go to https://docs.google.com
2. Click **Blank**
3. Rename it:
   ```
   DISC Introduction Reflection Survey — Master Responses
   ```
4. Move it into your **Pharming Kristy DISC Reflection Survey** folder on Google Drive (same folder as the Sheet)

---

## PART 2 — Upload the logo to Google Drive

1. On your computer, open:
   ```
   Cursor/Pharming Kristy DISC Reflection Survey /DISCconnector_Brand_Header.png
   ```
2. Upload that file into the **same Google Drive folder** as the Doc and Sheet

The script looks for `DISCconnector_Brand_Header.png` in that folder.

---

## PART 3 — Copy the Doc ID

Open the Google Doc. The URL looks like:
```
https://docs.google.com/document/d/XXXXXXXXXXXXXXXX/edit
```
Copy the **XXXXXXXXXXXXXXXX** part.

---

## PART 4 — Update Apps Script

1. Open Kristy’s **Google Sheet** → **Extensions → Apps Script**
2. Select **all** code → delete → paste **`DISC-Introduction-Reflection-Survey-AppsScript.gs`**
3. Find this line near the top:
   ```javascript
   DOC_ID: 'PASTE_DOC_ID_HERE',
   ```
4. Replace `PASTE_DOC_ID_HERE` with your Doc ID (keep the quotes)
5. **Save** (Cmd + S)
6. Confirm you see **`kristy-doc-v5`** if you search the file (Cmd + F)

---

## PART 5 — Apply formatting (logo, fonts, sizes)

**Important:** `setupAll` only runs once. To fix logo position/size and fonts on an existing doc, use:

1. In the function dropdown, choose **`applyAllDocFormatting`**
2. Click **Run** → authorize if asked
3. Open your Google Doc — you should see:
   - **DISC-Connector logo** — larger, **top left** in the header
   - **Title** — 16pt Poppins, purple, **not bold**
   - **Body text** — 12pt Poppins
   - Footer with copyright and disclaimer

If this is a brand-new empty doc, you can run **`setupAll`** instead (then run **`applyAllDocFormatting`** anyway to confirm styles).

---

## PART 6 — Deploy

1. **Deploy → Manage deployments → Edit → New version → Deploy**
2. Submit a **test survey**
3. Check the **Google Doc** — newest response should appear at the top

---

## Checklist

- [ ] Doc created and in Kristy folder on Drive
- [ ] `DISCconnector_Brand_Header.png` uploaded to same folder
- [ ] `DOC_ID` pasted in Apps Script
- [ ] `setupAll` ran successfully
- [ ] New deployment published
- [ ] Test response appears in Doc (newest first)
