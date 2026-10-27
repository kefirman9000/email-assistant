# Email Assistant

A simple Windows app that reads your unread work email in Outlook, suggests replies to the easy ones using an AI that runs entirely on your own computer, and lets you review, edit and approve each reply before it's sent.

**Nothing is ever sent without you clicking "Approve & send".**

## What you need

- A Windows PC with the **Outlook desktop app** open and signed in
- **Python** – <https://www.python.org/downloads/> (tick *"Add python.exe to PATH"* while installing)
- **Ollama** – <https://ollama.com/download> (the AI runs locally; your email never leaves your PC)

## First-time setup

1. Download the app: <https://github.com/kefirman9000/email-assistant/archive/refs/heads/stable.zip>. Unzip it somewhere like your Documents folder.
2. Double-click **Setup.bat** and wait for "All done!" (the AI download can take a while).
3. Double-click **Start Email Assistant**.
4. On the **Settings** tab, fill in your name, check the text that goes before and after every reply, and click **Save settings**.

Tip: right-click *Start Email Assistant* → *Send to* → *Desktop (create shortcut)*.

## Everyday use

1. Pick a time range under **Look at emails from** (past 30 minutes, past hour, today, past week...). Tick **Include emails I've already read** to look at read ones too. Then click **Check emails**.
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

Work happens on `main`. Whatever is on the **`stable`** branch is what users get.

1. Make your changes on `main`.
2. Edit `version.json`: raise `version` (e.g. `"1.1.0"`) and write `notes` in plain English. The notes are shown to the user as "What's new".
3. Commit, then publish by updating `stable`:
   ```
   git push origin main
   git push origin main:stable
   ```

Every copy of the app will offer the new version the next time it starts. A version tag (`v1.1.0`) or a GitHub release also works; the app picks whichever source has the highest version. The repository must be **public** so the app can download releases without a GitHub login.

Files that are never committed or overwritten by updates: `settings.json`, `assistant_data.json`, `.installed_version`, `backup/`.

## Settings reference

| Setting | What it does |
|---|---|
| Text before / after every reply | Greeting and sign-off wrapped around every suggestion. `{first_name}` = sender's first name, `{my_name}` = your name |
| Rules for the assistant | Extra do's and don'ts for the AI |
| Never suggest replies to these senders | Addresses or partial matches (e.g. `@vendor.com`) to always skip |
| Check automatically every (minutes) | `0` = only when you click the button |
| AI model name | Any model pulled in Ollama (default `llama3.1:8b`) |
