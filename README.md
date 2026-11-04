# Let's Wrassle Leads

Wholesale lead outreach for **Let's Wrassle Apparel Co.**, run entirely from one Google Sheet.

- Every Monday, a Claude scheduled task researches 10–15 new leads (wrestling clubs, academies and gear shops) and writes a personal pitch email for each one.
- The leads show up in the **Leads** tab of your Google Sheet.
- In the **Decision** column, pick **Approve** and the email sends from `support@letswrassleapparel.com`. Pick **Deny** and the lead is moved out of your queue.
- Leads with no reply after 5 days get one follow-up, which comes back to the queue for you to approve.

You review leads in a **phone web app** (`apps-script/Index.html`), which you can add to your iPhone home screen. You can also use the sheet directly. Both stay in sync because the app reads and writes the same tabs.

## Phone app

The web app is served by the same Apps Script project, so it runs as you, sends from your Gmail, and only you can open it.

- **Deploy:** in the Apps Script editor, go to **Deploy → New deployment → Web app**. Set *Execute as* to **Me** and *Who has access* to **Only myself**, then click **Deploy** and copy the `/exec` URL.
- **Home screen:** open that URL in **Safari** on your iPhone, tap **Share**, then **Add to Home Screen**.
- **Using it:**
  - **Approve & send** takes two taps, and the email goes out right away.
  - **Deny** removes the lead, with an Undo.
  - **Edit:** you can change the email, subject or message before you send.
  - **Instagram:** leads with a handle have a **Copy DM** button.
  - **Sent and Denied tabs** show history. You can move a denied lead back to review.
- **Updating:** after you change the code, use **Deploy → Manage deployments → Edit → Version: New version**. The URL stays the same.

```
Claude (weekly) ──CSV──▶ Drive folder "Let's Wrassle – Incoming Leads"
                                  │  (imported hourly, duplicates skipped)
                                  ▼
                       Google Sheet  ▸ Leads ▸ Decision = Approve ──▶ Gmail send ▸ Sent tab
                                              Decision = Deny    ──▶ Denied tab
                       (daily) no reply after 5 days ──▶ follow-up row back in Leads
```

## Tabs

| Tab | What it is |
|---|---|
| **Leads** | Your review queue. Edit **Email**, **Subject** or **Body** if you want, then set **Decision**. |
| **Sent** | Everything emailed, with the date and Gmail thread. Status becomes **Replied** when they answer. |
| **Denied** | Leads you said no to. They're kept so they're never suggested again. |
| **Settings** | From address, sender name, signature, follow-up days, daily send cap, incoming folder. |

## One-time setup (about 10 minutes)

> **Which Google account?** Use the Gmail account that has `support@letswrassleapparel.com` set up as a "Send mail as" address (`athletemindsetnetwork@gmail.com`). The script sends as whoever installs it. The incoming-leads folder is already shared with that account.

1. **Create the sheet.** Signed in as that account, go to [sheets.new](https://sheets.new) and name it **Let's Wrassle Leads**.
2. **Add the script.** Go to **Extensions → Apps Script**. Delete what's in `Code.gs` and paste in [`apps-script/Code.gs`](apps-script/Code.gs). Add a script file named `WebApp` with [`WebApp.gs`](apps-script/WebApp.gs), and an HTML file named `Index` with [`Index.html`](apps-script/Index.html).
   Then click the gear (**Project Settings**), check **Show "appsscript.json" manifest file in editor**, and replace that file with [`apps-script/appsscript.json`](apps-script/appsscript.json). Save.
3. **Run setup.** Reload the sheet. A **Let's Wrassle** menu appears. Click **Let's Wrassle → Run setup (first time)** and approve the Google permissions it asks for. You'll see an "unverified app" screen because it's your own script: click **Advanced → Go to project**.
   Setup creates the tabs, adds the Approve/Deny dropdown, installs the hourly import and daily follow-up checks, and imports your first 10 leads.
4. **Add your mailing address.** In **Settings → Signature**, replace `[ADD YOUR MAILING ADDRESS OR PO BOX HERE]`. CAN-SPAM requires a physical address in commercial email.
5. **Share the sheet with your Claude account.** Click Share, add `nathansmith2550@gmail.com` as a **Viewer**, and pick Viewer, not Editor. The weekly run reads the sheet so it never suggests a lead twice.

### Alternative: deploy with `clasp`
```bash
npm i -g @google/clasp && clasp login
cp .clasp.json.example .clasp.json   # paste your script ID (Apps Script → Project Settings)
clasp push
```

## Daily use

- **Approve:** the email sends right away, and the row moves to **Sent**.
- **Deny:** the row moves to **Denied**.
- **No email yet?** The Status says **Find contact** or **DM only**. Type an address into **Email**, then Approve. **DM draft** has a ready-to-paste Instagram message.
- **Something went wrong?** The Decision resets and **Status** shows a ⚠️ message. For example, an invalid email, a missing alias or the daily cap.
- **Menu shortcuts:** *Import new leads now*, *Check for follow-ups now*, and *Process Approve/Deny now*. The last one catches any decision the edit trigger missed.

## Notes and limits

- Consumer Gmail allows about **100 recipients a day** from Apps Script. Settings caps sends at 50 a day.
- Gmail signatures aren't available to scripts, so the signature comes from **Settings → Signature**.
- Follow-ups go out as a new message with the same `Re:` subject, so Gmail threads them with the original. Each lead gets at most one.
- Only use publicly listed business contacts. Every pitch ends with an easy opt-out line.

## Files

| Path | Purpose |
|---|---|
| `apps-script/Code.gs` | The whole automation: setup, Approve/Deny, CSV import, follow-ups |
| `apps-script/WebApp.gs` | The phone app's server side (list, approve, deny, restore) |
| `apps-script/Index.html` | The phone app's interface |
| `apps-script/appsscript.json` | Apps Script manifest (time zone, scopes) |
| `starter-leads.csv` | The first 10 researched leads (already in the Drive incoming folder) |
| `claude/weekly-task-prompt.md` | The prompt the Claude scheduled task runs each Monday |
| `.clasp.json.example` | For deploying with clasp |
