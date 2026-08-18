# AI Wine Agent

*A prototype built for Brooklyn Wine Club.*

## What this is

This is a tool that lets someone type a plain-English request — like "red wines under $50"
or "something great with steak" — into a spreadsheet of wine data, and have it instantly
filter and sort to match. No formulas, no digging through dropdown menus, no memorizing
which column holds what. Type what you're looking for, and it does the rest.

Right now it lives inside Excel as an add-in: open a wine list in Excel, click a button in
the sidebar, type a request, and the list filters and sorts itself.

It also works on wine lists that aren't laid out the same way — different column names,
different formatting — so it isn't tied to one specific spreadsheet template.

## Try it

If someone sent you this to try out, see:
- **[MAC_INSTALL.md](./MAC_INSTALL.md)** — setup steps for Mac
- **[WINDOWS_INSTALL.md](./WINDOWS_INSTALL.md)** — setup steps for Windows

Things to try typing once it's running:
- `red wines under $2,000`
- `something great with steak`
- `an argentine malbec, best rated first`
- `sparkling wines for a celebration`
- `wines not from the US`
- `something budget-friendly for a casual weeknight`

## What it can do today

- Understands wine **type**, **country**, **region**, **grape**, **price**, **rating**, and **vintage**
- Understands food pairings and occasions (steak, seafood, celebrations, budget-friendly, etc.) and matches them to sensible wine types/grapes
- Understands "not" — "not from the US," "excluding sparkling," "besides Malbec" all work as exclusions, not just inclusions
- Works on pretty much any wine spreadsheet — if the sheet isn't already formatted as a table, it sets that up automatically, and it recognizes a wide range of column names on its own (`Bottle`, `Retail $`, `Critic Score`, `Varietal`, and more)
- Recognizes roughly 30 wine-producing countries and 7 wine styles (red, white, rosé, sparkling, dessert, fortified, orange) by name
- Warns you if a price or rating column is stored as text (like `"$45.00"`) instead of a real number, since that can quietly break filtering
- Never rewrites or reformats your data — it only shows/hides and reorders rows, exactly like using Excel's own filter dropdowns by hand

## Where this could go

Right now, this only works inside a spreadsheet — someone needs the wine list open in
Excel for it to do anything. But the part of it that actually understands wine (what
"steak wine" means, which countries and styles it recognizes, how to read a plain-English
request) doesn't know or care that it's running inside Excel. That's just where it happens
to live today, because a spreadsheet was the fastest way to build and test a working
prototype.

That same matching logic could just as easily power a customer- or staff-facing experience
outside of Excel entirely:

- **A page on Brooklyn Wine Club's website** where a customer describes what they're
  looking for — "something for a dinner party, not too expensive" — and gets real
  suggestions pulled from current inventory, instead of browsing a long list.
- **A tool for staff on the floor or at checkout**, so anyone working that day can quickly
  answer "what do we have like this?" or "what pairs with X?" without needing to be
  comfortable in a spreadsheet.
- **A place for customers to save notes on wines they've tried** ("loved this, too tannic
  for my taste," "buy again for the holidays") and get suggestions based on what they've
  actually liked before — turning this from a one-time lookup into something that
  remembers a person's taste over time.
- **Recommendations built into a newsletter or loyalty program**, surfacing new arrivals
  based on someone's past purchases or saved notes.

None of that means starting over. The logic for what "steak wine" means, which countries
and styles and pairings it recognizes, and how it reads a plain-English request is already
written and tested — it would just need a different front door than an Excel sidebar.
Getting there is really two things: building a website or app interface in place of the
Excel task pane, and adding somewhere to actually store people's notes and preferences over
time, since a spreadsheet isn't built for that but a small database is a natural fit.

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

### How it recognizes any wine spreadsheet

Older versions of this add-in required the sheet to already be an Excel Table and used
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
   the sheet, so a sheet using its own vocabulary for a value still works even if that exact
   word isn't in the dictionary yet.
4. **Negation is handled explicitly.** "Not," "excluding," "besides," "other than," and
   similar cues in front of a value flip that match from an inclusion to an exclusion,
   scoped to that specific value — see `isNegated` in `parser.ts`.

If a column genuinely can't be found (e.g. there's no price column at all), the add-in just
skips that part of the prompt rather than failing — check the status message after clicking
**Apply** to see exactly what it matched.

### Requirements

- Office/Excel with the modern Table API (Microsoft 365 / Excel 2021+). No manual `Ctrl+T`
  step required — see above.
