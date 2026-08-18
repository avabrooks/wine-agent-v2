# Installing the AI Wine Agent in Excel (Mac)

This adds the wine-sorting assistant to your copy of Excel. You only need to do this once.

## What you'll need
- A Mac with Excel (the Microsoft 365 desktop version, not Excel Online)
- The `manifest.xml` file 

## Steps

1. **Save the `manifest.xml` file** somewhere you can find it easily, like your Desktop.

2. **Open Excel** and open any workbook (a blank one is fine).

3. Go to the **Insert** tab in the ribbon.

4. Click **Add-ins** → **My Add-ins**.
   *(On some versions this is a single "Add-ins" button either way, look for "My Add-ins" in the menu that appears.)*

5. In the dialog that opens, look for a small gear/settings icon or a link near the top called **Upload My Add-in**. Click it.

6. Browse to wherever you saved `manifest.xml`, select it, and click **Upload**.

7. The add-in should now appear look for **"AI Wine Agent"** (or a **Show Task Pane** button) in the **Home** tab of the ribbon, usually on the far right.

8. Click it to open the task pane, type what you're looking for and hit **Apply**.

## Using it

- Type a plain-English description of what you want and click **Apply** — it filters and sorts your wine table right in the sheet.
- Click **Clear Filters** any time to bring back the full list.
- Your spreadsheet needs to have your wine data set up as a proper Excel **Table** (with the little filter arrows in the header row) for this to work — if you're starting from scratch, select your data and press **⌘ + T** to make it a table first.

## If it stops showing up

Occasionally Excel's cache forgets about sideloaded add-ins (this can happen after an Excel update, or just over time). If that happens, just repeat steps 3–6 above with the same `manifest.xml` file — it takes 30 seconds.

## Troubleshooting

- **Nothing happens when I click Upload My Add-in** — make sure you selected the actual `manifest.xml` file, not a folder.
- **I get an error about the add-in failing to load** — check your internet connection; the add-in loads its interface from the web, so it needs to be online.
- **The task pane is blank or looks broken** — try closing and reopening the workbook, or restarting Excel.