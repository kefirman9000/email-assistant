# Email Assistant

A simple app that reads your work email, suggests replies to the easy ones using an AI that runs entirely on your own computer, and lets you review, edit and approve each reply before it's sent.

It can connect two ways (choose under **Settings → Connection**):

| | **Microsoft account** (recommended) | **Classic Outlook on this PC** |
|---|---|---|
| Works with | New Outlook, classic Outlook, Outlook on the web, Mac | Classic Outlook desktop app on Windows only |
| Setup | One-time IT approval, then sign in | Nothing extra, just have classic Outlook open |
| Outlook must be open | No | Yes |

**Nothing is ever sent without you clicking "Approve & send".**

## What you need

- A work Microsoft 365 email account. For the **Microsoft account** connection, IT approves the app once (see [IT setup](#it-setup)). For **Classic Outlook on this PC**, you just need the classic Outlook desktop app.
- **Python** – <https://www.python.org/downloads/> (tick *"Add python.exe to PATH"* while installing)
- **Ollama** – <https://ollama.com/download> (the AI runs locally; your email is never sent to an AI company)

## First-time setup

1. Download the app: <https://github.com/kefirman9000/email-assistant/archive/refs/heads/stable.zip>. Unzip it somewhere like your Documents folder.
2. Double-click **Setup.bat** and wait for "All done!" (the AI download can take a while).
3. Double-click **Start Email Assistant**.
4. On the **Settings** tab:
   - Fill in your name.
   - Under **Connection**, pick how to reach your email:
     - **Microsoft account**: enter the two codes from IT and click **Sign in**. A browser window opens; sign in with your work account.
     - **Classic Outlook on this PC**: nothing else to do. Keep Outlook open.
   - Check your greeting and sign-off under **Reply template**, then click **Save settings**.
5. Click **Test connection** to make sure everything works.

Tip: right-click *Start Email Assistant* → *Send to* → *Desktop (create shortcut)*.

## Everyday use

1. Pick a time range under **Look at emails from** (past 30 minutes, past hour, today, past week...). Tick **Include emails I've already read** to look at read ones too. Then click **Check emails**.
2. The **Inbox** tab (the start screen) lists every email from your checks, with its urgency and what the assistant decided. You can **View email** right in the app, **Open in Outlook**, or **Write a reply with AI**, even for emails the assistant skipped.
3. On **Suggested replies**, for each email you can:
   - Click in the white box to **edit the reply** (the grey greeting and sign-off are added automatically)
   - **Approve & send** (you'll see the full reply one more time before it goes)
   - **Rewrite with AI**, e.g. "more formal" or "say yes but suggest 3pm"
   - **Save as Outlook draft** to finish it in Outlook
   - **Move to 'needs attention'** or **Don't reply**
4. **Needs your attention** lists emails the assistant couldn't safely answer. Click **Have AI draft a reply** and tell it what to say, or **Start a reply in Outlook Drafts**.

## Updates

**Updates are only installed when you accept them.** When the app starts, it quietly checks for a new version. If there is one, an orange **Update available** button appears at the top. Nothing else happens until you click it, read what's new, and choose **Yes**. Only then does it download the update, keep a backup of the old version in the `backup` folder, install, and restart. Your settings, sign-in and pending replies are kept. You can also check any time from **Settings → Check for updates now**, or turn off the startup check there.

---

## IT setup

*Only needed for the **Microsoft account** connection. For the company's Microsoft 365 / Entra ID administrator. Takes about 5 minutes.*

**What the app is:** a desktop program that runs on the user's PC. It signs in as the user (delegated permissions only), reads and replies to **that user's own mailbox**, and runs its AI model locally through Ollama. There is no server, and no mail content is sent anywhere except Microsoft. Source code: <https://github.com/kefirman9000/email-assistant>.

1. In the **Microsoft Entra admin center**, go to **Identity → Applications → App registrations → New registration**.
   - **Name:** Email Assistant
   - **Supported account types:** Accounts in this organizational directory only (single tenant)
   - **Redirect URI:** platform **Public client/native (mobile & desktop)**, value `http://localhost`
2. Open the new app's **Authentication** page and set **Allow public client flows** to **Yes**. Save.
3. On **API permissions**, choose **Add a permission → Microsoft Graph → Delegated permissions** and add:
   - `User.Read` – see who is signed in
   - `Mail.ReadWrite` – read the inbox, create reply drafts, mark emails read and tag them "AI Replied"
   - `Mail.Send` – send the replies the user approves
   - `offline_access` – stay signed in between sessions
4. Click **Grant admin consent** for your organization (or let users consent themselves, if your policy allows it).
5. Give the user the **Application (client) ID** and **Directory (tenant) ID** from the app's **Overview** page. They enter these under **Settings → Microsoft account**.

Conditional Access and MFA apply to sign-in as normal. Users can revoke access at any time from <https://myapps.microsoft.com>, and admins can remove the app registration.

---

## For the maintainer

### Publishing an update

Work happens on `main`. Whatever is on the **`stable`** branch is what users get.

1. Make your changes on `main`.
2. Edit `version.json`: raise `version` (e.g. `"2.1.0"`) and write `notes` in plain English. The notes are shown to the user as "What's new".
3. Commit, then publish by updating `stable`:
   ```
   git push origin main
   git push origin main:stable
   ```

Every copy of the app will offer the new version the next time it starts. A version tag (`v2.1.0`) or a GitHub release also works; the app picks whichever source has the highest version. The repository must be **public** so the app can download updates without a GitHub login.

### Skipping the client ID step for users

If everyone uses the same app registration, put its client ID in `DEFAULT_CLIENT_ID` near the top of `email_assistant.py` (and the tenant ID in `DEFAULT_TENANT`). The App ID box then disappears from Settings.

Files that are never committed or overwritten by updates: `settings.json`, `assistant_data.json`, `signin_cache.bin`, `.installed_version`, `backup/`. The sign-in cache is encrypted with Windows DPAPI (or the macOS Keychain).

## Settings reference

| Setting | What it does |
|---|---|
| Connection | **Microsoft account** (App ID and Directory ID from IT, then Sign in) or **Classic Outlook on this PC**. Test connection works for both |
| Reply template | Greeting and sign-off wrapped around every suggestion. `{first_name}` = sender's first name, `{my_name}` = your name. Live preview included |
| Rules for the assistant | Extra do's and don'ts for the AI |
| Never suggest replies to these senders | Addresses or partial matches (e.g. `@vendor.com`) to always skip |
| Auto-check every (minutes) | `0` = only when you click the button |
| AI model name | Any model pulled in Ollama (default `llama3.1:8b`) |
