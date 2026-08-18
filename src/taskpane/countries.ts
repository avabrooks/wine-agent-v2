// Canonical wine-producing countries mapped to words a user might type in a
// prompt (adjectival forms, nicknames, abbreviations). Matched against the
// prompt in parser.ts; if a country doesn't appear here, the parser still
// falls back to matching whatever text literally appears in the sheet's
// Country column, so unlisted countries aren't a hard blocker — but adding
// them here lets people use natural phrasing like "an Argentine malbec"
// instead of typing the exact cell value.
export const COUNTRY_SYNONYMS: Record<string, string[]> = {
  France: ["france", "french"],
  Italy: ["italy", "italian"],
  Spain: ["spain", "spanish"],
  Portugal: ["portugal", "portuguese"],
  Germany: ["germany", "german"],
  Chile: ["chile", "chilean"],
  Argentina: ["argentina", "argentinian", "argentine"],

  "United States": [
    "united states",
    "usa",
    "u.s.a.",
    "us",
    "u.s.",
    "america",
    "american",
    "domestic",
    "california",
    "californian",
    "oregon",
    "washington state"
  ],
  Australia: ["australia", "australian", "aussie"],
  "New Zealand": ["new zealand", "nz", "kiwi"],
  "South Africa": ["south africa", "south african"],
  Austria: ["austria", "austrian"],
  Greece: ["greece", "greek"],
  Hungary: ["hungary", "hungarian"],
  Georgia: ["georgia", "georgian"],
  Israel: ["israel", "israeli"],
  Canada: ["canada", "canadian"],
  "United Kingdom": ["united kingdom", "uk", "england", "english", "britain", "british"],
  Uruguay: ["uruguay", "uruguayan"],
  Croatia: ["croatia", "croatian"],
  Slovenia: ["slovenia", "slovenian"],
  Romania: ["romania", "romanian"],
  Bulgaria: ["bulgaria", "bulgarian"],
  Moldova: ["moldova", "moldovan"],
  Lebanon: ["lebanon", "lebanese"],
  Morocco: ["morocco", "moroccan"],
  Brazil: ["brazil", "brazilian"],
  Mexico: ["mexico", "mexican"],
  Japan: ["japan", "japanese"],
  China: ["china", "chinese"],
  Switzerland: ["switzerland", "swiss"],
  Slovakia: ["slovakia", "slovakian", "slovak"],
  Serbia: ["serbia", "serbian"],
  Turkey: ["turkey", "turkish"],
  Armenia: ["armenia", "armenian"]
};
