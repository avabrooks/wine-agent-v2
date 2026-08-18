// Canonical wine "Type" values mapped to words a user might type in a prompt.
// e.g. typing "bubbly" or "prosecco" should match the "Sparkling" column value.
// These are matched two ways (see parser.ts): against the user's prompt, and
// against whatever text actually appears in the sheet's Type column — so a
// sheet that spells it "Red Wine" or "Sparkling Wine" still lines up.
export const TYPE_SYNONYMS: Record<string, string[]> = {
  red: ["red", "red wine"],
  white: ["white", "white wine"],
  rose: ["rose", "rosé", "rosado", "rosato", "pink", "blush"],
  sparkling: [
    "sparkling",
    "champagne",
    "prosecco",
    "cava",
    "bubbly",
    "crémant",
    "cremant",
    "sparkling wine"
  ],
  dessert: ["dessert", "sweet", "late harvest", "ice wine", "icewine", "dessert wine"],
  fortified: ["fortified", "port", "sherry", "madeira", "marsala", "vermouth", "fortified wine"],
  orange: ["orange", "orange wine", "skin contact", "skin-contact", "amber wine"]
};

// Canonical spreadsheet fields mapped to the header names people actually use.
// Lets the add-in work with sheets titled "Wine", "Label", "Varietal", "Score",
// "Bottle", "Retail $", "Critic Score", etc. instead of requiring an exact
// "Wine Name" / "Grape" / "Rating" header. Matching also falls back to
// substring and fuzzy-spelling checks (see detectColumns in parser.ts), so
// this list only needs to cover common phrasing, not every possible header.
export const HEADER_SYNONYMS: Record<string, string[]> = {
  name: [
    "wine name",
    "name",
    "wine",
    "label",
    "product",
    "bottle",
    "producer",
    "winery",
    "item",
    "wine/producer",
    "description"
  ],
  country: ["country", "origin", "nation", "country of origin"],
  region: ["region", "appellation", "area", "subregion", "sub-region", "region/appellation"],
  type: ["type", "color", "colour", "wine type", "style", "category"],
  grape: ["grape", "varietal", "variety", "grape variety", "cepage", "cépage"],
  vintage: ["vintage", "year", "vintage year"],
  price: [
    "price",
    "cost",
    "$",
    "retail",
    "retail price",
    "bottle price",
    "unit price",
    "price per bottle",
    "$/btl",
    "cost per bottle"
  ],
  rating: [
    "rating",
    "score",
    "points",
    "critic score",
    "wine score",
    "avg rating",
    "average rating",
    "rated"
  ]
};
