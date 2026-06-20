# Add client email copy (Google, from surveys@discconnector.com)

**Time:** ~15 minutes  
**Account:** apgblue3@gmail.com

---

## Step 1 — Gmail “Send mail as” (do this first)

1. Open Gmail as **apgblue3@gmail.com**
2. **Settings (gear) → See all settings → Accounts**
3. Under **Send mail as**, click **Add another email address**
4. Enter:
   - Name: `DISC Connector`
   - Email: `surveys@discconnector.com`
5. Complete verification (Google will email a code or use your domain/DNS if Workspace)

Until this is verified, client copies **will fail**.

---

## Step 2 — Paste Apps Script into Kristy’s Sheet

1. Open **`DISC Introduction: Reflection Survey.gsheet`** (in this folder)
2. **Extensions → Apps Script**
3. Delete any old code in the editor
4. Paste all of **`DISC-Introduction-Reflection-Survey-AppsScript.gs`**
5. Save. Name the project: `Kristy Reflection Survey`

---

## Step 3 — Run setup once

1. In Apps Script, select function **`setupSheet`**
2. Click **Run** → authorize when asked
3. Check the Sheet has column headers (Timestamp, Name, Organization, Email, Q1–Q15)

---

## Step 4 — Deploy web app

1. **Deploy → New deployment → Web app**
2. Execute as: **Me**
3. Who has access: **Anyone**
4. Click **Deploy** and copy the **Web app URL** (ends in `/exec`)

---

## Step 5 — Update the HTML

1. Open **`DISC-Introduction-Reflection-Survey.html`**
2. Find: `var APPS_SCRIPT_URL = '...'`
3. Replace with your new Web app URL from Step 4
4. Save and re-upload to your website if it’s hosted online

---

## What happens after setup

| Email | To | From (what they see) |
|-------|-----|----------------------|
| Admin copy | surveys@discconnector.com | surveys@discconnector.com |
| Client copy | email they typed on the form | surveys@discconnector.com |

Responses still save to the Google Sheet.

---

## Quick test

1. Submit a test with your own email
2. Check inbox for **two** emails from surveys@discconnector.com (admin + your copy)
3. Check the Sheet for a new row

If client copy fails: Gmail **Send mail as** is not verified yet.
