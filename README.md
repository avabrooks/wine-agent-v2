# AI Wine Agent

An Excel add-in for **Brooklyn Wine Club** that lets you sort, search, and pair a wine
collection using plain-English prompts, right inside your spreadsheet — no matter how
that spreadsheet is laid out.

Type something like:
- `red wines under $2000`
- `steak wines under $40000`
- `something budget-friendly for a casual weeknight`
- `an argentine malbec, best rated first`

...and it filters and sorts your data accordingly, using Excel's own native filter and sort.

## Features

- Understands wine **type**, **country**, **region**, **grape**, **price**, **rating**, and **vintage**
- Understands food pairings and occasions (steak, seafood, celebrations, budget-friendly, etc.) and maps them to sensible wine types/grapes
- Works on **any wine spreadsheet** — if the sheet isn't already an Excel Table, the add-in creates one automatically from the data it finds
- Recognizes a wide range of column headers out of the box (`Bottle`, `Retail $`, `Critic Score`, `Varietal`, ...) via synonym lists plus fuzzy/substring matching, so it isn't locked to one exact template
- Recognizes ~30 wine-producing countries and 7 wine styles (red, white, rosé, sparkling, dessert, fortified, orange) by name
- Flags when a price/rating column is stored as text (e.g. `"$45.00"`) instead of numbers, since that can throw off Excel's native filter/sort
- Filters and sorts using Excel's native Table API — it never rewrites a cell's value or formatting
- "Clear Filters" instantly restores the full list

## For users: installing this add-in

If someone sent you a `manifest.xml` file to try this out, see **[MAC_INSTALL.md](./MAC_INSTALL.md)** for step-by-step setup instructions (Mac).

## For developers

### Project structure

```
src/
  taskpane/
    taskpane.html   — the task pane UI
    taskpane.css    — styling
    taskpane.ts     — wires up the Apply/Clear buttons
    excel.ts        — talks to the Excel Table API (filter, sort, clear, auto-create table)
    parser.ts       — turns a prompt into a filter/sort plan; column detection lives here
    constants.ts    — wine type + column header synonym dictionaries
    countries.ts    — country name synonym dictionary
    pairings.ts     — food/occasion/style pairing dictionary
manifest.xml              — the Office Add-in manifest (local dev / localhost)
manifest-production.xml   — the Office Add-in manifest (GitHub Pages deployment)
```

### Running locally

```
npm install
npm start
```

This launches a local dev server and sideloads the add-in into Excel for testing.

### Building for deployment

```
npm run build
```

Produces the compiled files used for hosting (see below).

### Hosting / deployment

This add-in is hosted via GitHub Pages. `manifest-production.xml`'s `SourceLocation`,
`IconUrl`, and `AppDomain` entries point at the GitHub Pages URL rather than localhost,
so anyone with that manifest file can sideload it without running the project themselves.
Update the URLs in `webpack.config.js` (`urlProd`) and both manifest files if you deploy
under a different repo name or custom domain.

## How it recognizes any wine spreadsheet

Older versions of this add-in required your sheet to already be an Excel Table and used
an exact-match header list, which meant it only worked reliably on one specific layout.
This version relaxes both constraints:

1. **No Table? No problem.** If the active sheet isn't already an Excel Table, the add-in
   finds the used range, skips any leading blank/title rows, and creates a Table from what's
   left — using the first surviving row as headers.
2. **Column detection is fuzzy.** Headers are matched against `HEADER_SYNONYMS` in three
   passes: exact match, then substring match (`"Retail Price (USD)"` still matches `price`),
   then a fuzzy spelling match for typos or near-misses.
3. **Values, not just headers, are fuzzy too.** Type/country matching checks both the synonym
   dictionaries in `constants.ts` / `countries.ts` *and* whatever text actually appears in
   your sheet, so a sheet using its own vocabulary for a value still works even if that exact
   word isn't in the dictionary yet.

If a column genuinely can't be found (e.g. there's no price column at all), the add-in just
skips that part of the prompt rather than failing — check the status message after clicking
**Apply** to see exactly what it matched.

## Requirements

- Office/Excel with the modern Table API (Microsoft 365 / Excel 2021+). No manual `Ctrl+T`
  step required — see above.
