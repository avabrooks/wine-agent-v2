# Running / installing the AI Wine Agent in Excel (Windows)

There are two different things this covers:
- **"Try it right now / keep developing it"** → use `npm start` (Section A). This is what you want while testing changes before pushing to GitHub.
- **"Share a finished build with someone else"** → sideload the built `manifest.xml` by hand (Section B). This is the Windows equivalent of what `MAC_INSTALL.md` covers for Mac.

## What you'll need

- Windows with Excel (the Microsoft 365 desktop version, not Excel Online)
- [Node.js](https://nodejs.org) (LTS version) — check if it's already installed by opening **Command Prompt** and running `node -v`. If that errors, install Node.js first.

## A. Live dev testing with `npm start`

1. Unzip this project folder somewhere, e.g. `C:\Users\YourName\wine-agent-v2`.
2. Open **Command Prompt** (not PowerShell — see the note below if you'd rather use PowerShell) and `cd` into the folder:
   ```
   cd C:\Users\YourName\wine-agent-v2
   ```
3. Install dependencies:
   ```
   npm install
   ```
4. Start it:
   ```
   npm start
   ```

The first time you run this, it installs a local HTTPS development certificate so Excel can load the add-in securely — Windows will likely pop up a **security warning asking to trust a certificate**. Click **Yes**. It should then start a local server and **automatically open Excel with "AI Wine Agent" already sideloaded** in the Home tab — no manual upload needed for this path.

When you're done testing, run `npm stop` to un-sideload it and stop the server.

### If PowerShell blocks `npm`

PowerShell sometimes refuses to run npm's script wrapper with an error like *"running scripts is disabled on this system."* Either:
- Use **Command Prompt (cmd.exe)** instead of PowerShell for these commands, or
- Run PowerShell as Administrator once and execute `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, then retry.

### If Excel doesn't open, or the task pane looks blank/broken

- Make sure Excel was fully closed before running `npm start` the first time.
- Try `npm stop` then `npm start` again.
- Restart Excel.

## B. Manually sideloading a built version (sharing with someone else)

1. Build the production files:
   ```
   npm run build
   ```
   This produces a `dist` folder with the compiled add-in and `manifest-production.xml`, pointing at wherever this is hosted (GitHub Pages, once pushed).
2. Send the person `manifest-production.xml` (they don't need the rest of the project).
3. They open **Excel** → **Insert** tab → **Add-ins** → **My Add-ins**.
4. Look for a small gear/settings icon or a link near the top called **Upload My Add-in**. Click it.
5. Browse to the `manifest-production.xml` file they were sent, select it, and click **Upload**.
6. **"AI Wine Agent"** should now appear in the **Home** tab of the ribbon.

Note: this only works once the add-in's actual files (`taskpane.html`, etc.) are hosted somewhere reachable — i.e. after this repo is pushed and GitHub Pages is turned on, since `manifest-production.xml` points there rather than at your own computer.

## Using it

- Type a plain-English description of what you want and click **Apply** — it filters and sorts your wine table right in the sheet.
- Click **Clear Filters** any time to bring back the full list.
- Your spreadsheet doesn't need to already be an Excel Table — if it isn't one yet, the add-in creates one automatically from your data the first time you click Apply.
