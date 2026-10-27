# Email Assistant

A simple Windows app that reads your unread work email in Outlook, suggests replies to the easy ones using an AI that runs entirely on your own computer, and lets you review, edit and approve each reply before it's sent.

**Nothing is ever sent without you clicking "Approve & send".**

## What you need

- A Windows PC with the **Outlook desktop app** open and signed in
- **Python** – <https://www.python.org/downloads/> (tick *"Add python.exe to PATH"* while installing)
- **Ollama** – <https://ollama.com/download> (the AI runs locally; your email never leaves your PC)

## First-time setup

1. Download the latest version: on this page click **Releases** → newest release → **Source code (zip)**, then unzip it somewhere like your Documents folder.
2. Double-click **Setup.bat** and wait for "All done!" (the AI download can take a while).
3. Double-click **Start Email Assistant**.
4. On the **Settings** tab, fill in your name, check the text that goes before and after every reply, and click **Save settings**.

Tip: right-click *Start Email Assistant* → *Send to* → *Desktop (create shortcut)*.

## Everyday use

1. Click **Check for new emails**.
2. On **Suggested replies**, for each email you can:
   - **Open the original email** to read what was sent
   - Click in the white box to **edit the reply** (the grey text above and below is added automatically and can be changed in Settings)
   - **Approve & send** (you'll see the full reply one more time before it goes)
   - **Rewrite with AI**, e.g. "more formal" or "say yes but suggest 3pm"
   - **Edit in Outlook** to finish it in Outlook (also saved in Drafts)
   - **Move to 'needs attention'** or **Don't reply**
3. **Needs your attention** lists emails the assistant couldn't safely answer. Click **Have AI draft a reply** and tell it what to say, or **Reply myself in Outlook**.

## Updates

The app checks for updates when it starts (and from **Settings → Check for updates now**). When a new version is out, it shows what's new and asks before installing. It then downloads the update, keeps a backup of the old version in the `backup` folder, installs, and restarts. Your settings and pending replies are kept.

---

## For the maintainer: publishing an update

1. Make your changes and bump `APP_VERSION` near the top of `email_assistant.py` (e.g. `"1.1.0"`).
2. Commit and push to `main`.
3. Create a GitHub release whose tag matches the version, with notes in plain English (they're shown to the user):
   ```
   gh release create v1.1.0 --title "Version 1.1.0" --notes "Replies now keep your Outlook signature."
   ```
   or on GitHub: **Releases → Draft a new release**.

Every copy of the app will offer the update next time it starts. The repository must be **public** so the app can download releases without a GitHub login.

Files that are never committed or overwritten by updates: `settings.json`, `assistant_data.json`, `.installed_version`, `backup/`.

## Settings reference

| Setting | What it does |
|---|---|
| Text before / after every reply | Greeting and sign-off wrapped around every suggestion. `{first_name}` = sender's first name, `{my_name}` = your name |
| Rules for the assistant | Extra do's and don'ts for the AI |
| Never suggest replies to these senders | Addresses or partial matches (e.g. `@vendor.com`) to always skip |
| Look at emails from the last (hours) | How far back each check goes (unread mail only) |
| Check automatically every (minutes) | `0` = only when you click the button |
| AI model name | Any model pulled in Ollama (default `llama3.1:8b`) |
