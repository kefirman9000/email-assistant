# Email Assistant

A simple app that reads your work email, suggests replies to the easy ones using an AI that runs entirely on your own computer, and lets you review, edit and approve each reply before it's sent.

It can connect two ways (choose under **Settings → Connection**):

| | **Microsoft account** (recommended) | **Classic Outlook on this PC** |
|---|---|---|
| Works with | New Outlook, classic Outlook, Outlook on the web, Mac | Classic Outlook desktop app on Windows only |
| Setup | One-time IT approval, then sign in | Nothing extra, just have classic Outlook open |
| Outlook must be open | No | Yes |

**Nothing is ever sent without you clicking "Approve & send".**

## Screenshots

**Inbox**: every email from your checks, with its urgency and what the assistant decided.

![Inbox](docs/screenshots/inbox.png)

**Reply picker**: two AI replies side by side. Pick a vibe like *Friendly denial* to get two new options in that style.

![Reply picker](docs/screenshots/reply-picker.png)

**Suggested replies**: edit, choose whether to include your greeting and sign-off, use a quick reply, then approve and send.

![Suggested replies](docs/screenshots/suggested-replies.png)

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

1. Pick a time range under **Show emails from** (past 30 minutes, past hour, today, past week...). Tick **Include read emails** to look at read ones too. Then click **Check emails**. To check automatically, pick an interval under **Auto-refresh** (every 5, 10, 15 or 30 minutes, or every hour).
2. The **Inbox** tab (the start screen) lists every email from your checks, with its urgency and what the assistant decided. You can **View** an email right in the app, **Open in Outlook**, or **Write a reply with AI**, even for emails the assistant skipped.
3. On **Suggested replies**, for each email you can:
   - Click in the white box to **edit the reply** (the grey greeting and sign-off are added automatically)
   - **Approve & send** (you'll see the full reply one more time before it goes)
   - **Rewrite with AI**: shows **two different replies** side by side. Click a vibe (*Happier, More frank, Friendly acceptance, Friendly denial, More formal, More casual, Shorter, Apologetic, Ask for details*) for two new options in that style, or describe what you want (e.g. "say yes but suggest 3pm"). Edit either one, then click **Use this reply**
   - **Quick reply** to drop in a saved one-click reply (no AI), like "Thanks, got it!"
   - Untick **Include greeting and sign-off** to send just the reply text
   - **Save as draft** to finish it in Outlook
   - **Needs attention** to move it to the other list, or **Don't reply**
4. **Needs your attention** lists emails the assistant couldn't safely answer. Click **Draft a reply with AI** to choose between AI options the same way, or **Start my own reply**.

## Updates

**Updates are only installed when you accept them.** When the app starts, it quietly checks for a new version. If there is one, an orange **Update available** button appears at the top. Nothing else happens until you click it, read what's new, and choose **Yes**. Only then does it download the update, keep a backup of the old version in the `backup` folder, install, and restart. Your settings, sign-in and pending replies are kept. You can also check any time with the **Check for updates** button at the top of the window, or turn off the startup check under **Settings → Updates**.

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

## Settings reference

| Setting | What it does |
|---|---|
| Connection | **Microsoft account** (App ID and Directory ID from IT, then Sign in) or **Classic Outlook on this PC**. Test connection works for both |
| Reply template | Greeting and sign-off wrapped around every suggestion. `{first_name}` = sender's first name, `{my_name}` = your name. Live preview included |
| Rules for the assistant | Extra do's and don'ts for the AI |
| Never suggest replies to these senders | Addresses or partial matches (e.g. `@vendor.com`) to always skip |
| Quick replies | Your saved one-click replies. Each has its own **Greeting & sign-off** switch |
| Notifications | A Windows notification (and a flashing taskbar button) when a check finds new emails that need you. Option for urgent emails only, plus a **Send a test notification** button |
| Sending pace | Replies go out one at a time with a random pause between them, never within a few minutes of an email arriving, and up to an hourly limit, so they look hand-sent rather than automated |
| Appearance | **Match Windows**, **Light** or **Dark**. There's also a **Dark mode** switch at the top of the window |
| AI model name | Any model pulled in Ollama (default `llama3.1:8b`) |
