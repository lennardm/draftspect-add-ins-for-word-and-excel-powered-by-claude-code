# Installing Draftspect on Windows (no admin rights needed)

This guide sets up Draftspect for Word and Excel on a Windows PC **without
administrator rights**, using **OpenRouter** for the AI (so there is no Claude
sign-in to do — just one API key in a text file).

Everything installs into your own user folder. Nothing here touches
system-wide settings or needs an IT admin.

You will do five things:

1. Install Node.js into a local folder and put it on your PATH
2. Get the Draftspect files
3. Run `npm install` inside the Draftspect folder
4. Paste your OpenRouter key into a `.env` file, then start it
5. Install add-ins for word and excel

---

## What you need first

- **Windows 10 or 11.** (Windows 11 already includes the WebView2 runtime that
  Office add-ins use. On Windows 10 it's almost always present too.)
- **Microsoft Word and Excel** (desktop versions — Microsoft 365 or Office
  2021+). The web versions won't work.
- **An OpenRouter account + API key.** Create one at
  <https://openrouter.ai/keys>. Add a few dollars of credit to the account so
  requests aren't rejected.

---

## Step 1 — Install Node.js into a local folder

We'll use the **portable ZIP** of Node.js, which needs no installer and no
admin rights.

1. Go to **<https://nodejs.org/en/download>**.

2. Choose **Windows**, **x64**, and the **binary `.zip`** (NOT the `.msi`
   installer). Download it.

3. Open the downloaded `.zip`. Inside is a single folder named something like
   `node-v22.x.x-win-x64`.

4. Copy that inner folder into your user folder and **rename it to `nodejs`**,
   so the final path is exactly:
   
   ```
   C:\Users\<your-name>\nodejs
   ```
   
   (It should contain `node.exe` and `npm.cmd` directly inside it.)

### Put Node.js on your PATH (so `node` and `npm` work when you type them)

5. Open **Windows PowerShell** (press Start, type *PowerShell*, press Enter).

6. Copy-paste this **one** line and press Enter. It adds your local Node folder
   to the front of your **personal** PATH (no admin, doesn't affect anyone
   else). Run it only once:
   
   ```powershell
   [Environment]::SetEnvironmentVariable("Path", "$env:USERPROFILE\nodejs;" + [Environment]::GetEnvironmentVariable("Path","User"), "User")
   ```

7. **Close that PowerShell window and open a brand-new one** (PATH changes only
   apply to windows opened afterwards).

8. Check it worked:
   
   ```powershell
   where.exe node
   where.exe npm
   node -v
   npm -v
   ```
   
   - The **first** line from each `where.exe` must point at your
     `C:\Users\<your-name>\nodejs` folder. If a different path shows up first,
     your local folder isn't first on PATH — redo step 6 in a new window.
   - `node -v` and `npm -v` should print version numbers.
   
   > **Note on "`.com` files":** Windows decides which file to run from the
   > extensions in `PATHEXT` (`.COM`, then `.EXE`, then `.BAT`, then `.CMD`…).
   > Node ships `node.exe` and `npm.cmd` — there are **no** `.com` files and you
   > don't need to create any. As long as your `nodejs` folder is **first** on
   > PATH (verified by `where.exe` above), typing `node` and `npm` uses these
   > local copies.

---

## Step 2 — Get the Draftspect files

If you were given the project as a ZIP or a folder, just put it somewhere in
your user area, e.g. `C:\Users\<your-name>\Draftspect`, and skip ahead.

Otherwise, download it from GitHub (no git required):

1. Open the project's GitHub page in a browser.

2. Click the green **Code** button → **Download ZIP**.

3. Extract the ZIP into your user folder, e.g.:
   
   ```
   C:\Users\<your-name>\Draftspect
   ```
   
   This folder (the one containing `package.json`) is what we'll call the
   **Draftspect folder** from now on.

---

## Step 3 — Install Draftspect's components

1. Open a **new** PowerShell window.

2. Move into the Draftspect folder (adjust the path if you used a different
   location):
   
   ```powershell
   cd "$env:USERPROFILE\Draftspect"
   ```

3. Install everything:
   
   ```powershell
   npm install
   ```
   
   This downloads Draftspect's components, **including the bundled Claude Code
   engine** and the desktop tray app. It can take a few minutes and pulls
   ~100–200 MB the first time. All of it lands inside the Draftspect folder —
   nothing system-wide, no admin.
   
   > If your company uses a network proxy and the download fails, you may need
   > the proxy address from your IT team. That's the only step a corporate
   > network sometimes interferes with.

---

## Step 4 — Add your OpenRouter key and start it

1. Make your own settings file by copying the template (run this in the
   Draftspect folder):
   
   ```powershell
   copy .env.example .env
   ```

2. Open the new `.env` file in **Notepad**:
   
   ```powershell
   notepad .env
   ```

3. Find the line:
   
   ```
   OPENROUTER_API_KEY=
   ```
   
   and paste your key right after the `=` (no spaces, no quotes), e.g.:
   
   ```
   OPENROUTER_API_KEY=sk-or-v1-abc123...
   ```
   
   Save the file (Ctrl+S) and close Notepad. Everything else in the file is
   optional and can be left as-is.

4. Start Draftspect:
   
   ```powershell
   npm start
   ```
   
   A small **Draftspect** icon appears in the Windows system tray (bottom-right,
   you may need to click the `^` to show hidden icons). Leave this running —
   it's the engine.

5. **Confirm OpenRouter is active.** Right-click the tray icon → **Open logs**.
   Near the top you should see a line like:
   
   ```
   [daemon] OpenRouter routing ENABLED (dev toggle) — base=https://openrouter.ai/api ...
   ```
   
   If you see that, the key was read correctly.

---

## Step 5 — Install the add-in into Word and Excel

1. Right-click the **Draftspect tray icon** → **Install add-in in Word +
   Excel…**. (This registers the add-in just for your user account — no admin.)
2. **Fully close Word and Excel** if they're open, then reopen one of them.
3. In Word/Excel: **Insert** tab → **Add-ins** (or **My Add-ins**) → **Shared
   Folder** → choose **Draftspect**. The Draftspect panel opens on the right.
4. In the panel, pick a **workspace folder** (where your documents live), and
   start chatting.

---

## Everyday use after setup

- Make sure the **Draftspect tray app is running** (Step 5's `npm start`). If
  it isn't, open PowerShell, `cd` into the Draftspect folder, and run
  `npm start` again.
- Open Word/Excel and use the Draftspect panel as normal.

---

## If something goes wrong

| Problem                                           | Fix                                                                                                                                         |
| ------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| `npm` or `node` "not recognized"                  | You're in an **old** terminal. Close it, open a new one. Re-check Step 1.8.                                                                 |
| `where.exe node` shows the wrong folder first     | Your `nodejs` folder isn't first on PATH. Redo Step 1.6 in a fresh window.                                                                  |
| Log says usage limit / auth error                 | The OpenRouter key is missing, mistyped, or the account has no credit. Re-check `.env` (Step 5) and your OpenRouter balance.                |
| Log does **not** say "OpenRouter routing ENABLED" | The `.env` file is missing or the key line is blank. Confirm `.env` is in the Draftspect folder and the key is after `OPENROUTER_API_KEY=`. |
| Draftspect panel doesn't appear in Word/Excel     | Fully quit Word/Excel (not just close the document) and reopen. Re-run "Install add-in" from the tray if needed.                            |
| A model gives a "not found" error                 | Uncomment and edit the `OPENROUTER_MODEL_*` lines in `.env` to a model slug listed on OpenRouter, then restart from the tray.               |

For deeper troubleshooting and how the pieces fit together, see `README.md`.

---

## Appendix — using a Claude subscription instead of OpenRouter

OpenRouter is the easy, no-sign-in path. If you'd rather run on a personal
**Claude** subscription instead, leave `OPENROUTER_API_KEY` blank in `.env`
and follow the Claude Code sign-in instructions in `README.md`. (That path
uses your Claude account's OAuth; OpenRouter is not involved.)
