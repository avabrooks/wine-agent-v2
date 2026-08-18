// Knowledge base for prompts that describe a food, occasion, or mood rather than
// literal spreadsheet values (e.g. "something great with steak", "for a celebration").
// Each entry lists trigger keywords plus what it suggests: candidate Type/Grape
// values, a max price, or a minimum rating. Type/Grape suggestions only kick in
// when the prompt didn't already state a type or grape directly — see
// hadDirectType / hadDirectGrape in parser.ts — so an explicit "red wine" always
// wins over an inferred pairing. Add more entries here any time.
export interface PairingEntry {
  label: string;
  keywords: string[];
  types?: string[];
  grapes?: string[];
  excludeTypes?: string[];
  priceMax?: number;
  ratingMin?: number;
}

export const PAIRING_DICTIONARY: PairingEntry[] = [
  // --- Food pairings ---
  {
    label: "fish/seafood",
    keywords: ["white fish", "fish", "seafood", "sole", "cod", "halibut", "flounder", "trout"],
    types: ["White", "Sparkling"],
    grapes: ["Sauvignon Blanc", "Pinot Grigio", "Albarino", "Chardonnay", "Riesling"]
  },
  {
    label: "shellfish",
    keywords: ["shellfish", "shrimp", "crab", "lobster", "oysters", "scallops", "clams"],
    types: ["Sparkling", "White"],
    grapes: ["Chardonnay", "Albarino", "Riesling"]
  },
  {
    label: "sushi",
    keywords: ["sushi", "sashimi", "raw fish", "poke"],
    types: ["Sparkling", "White"],
    grapes: ["Riesling", "Sauvignon Blanc"]
  },
  {
    label: "salmon/fatty fish",
    keywords: ["salmon", "tuna", "fatty fish"],
    types: ["White", "Rose", "Red"],
    grapes: ["Pinot Noir", "Chardonnay"]
  },
  {
    label: "steak/red meat",
    keywords: ["steak", "red meat", "beef", "filet", "ribeye", "prime rib"],
    types: ["Red"],
    grapes: ["Cabernet Sauvignon", "Malbec", "Syrah", "Tempranillo", "Zinfandel"]
  },
  {
    label: "lamb",
    keywords: ["lamb"],
    types: ["Red"],
    grapes: ["Syrah", "Tempranillo", "Cabernet Sauvignon"]
  },
  {
    label: "pork",
    keywords: ["pork", "ham", "pork chop", "pork belly"],
    types: ["Red", "White"],
    grapes: ["Pinot Noir", "Tempranillo", "Riesling", "Chenin Blanc"]
  },
  {
    label: "chicken/poultry",
    keywords: ["chicken", "poultry", "roast chicken"],
    types: ["White", "Red"],
    grapes: ["Chardonnay", "Viognier", "Pinot Noir"]
  },
  {
    label: "turkey/holiday dinner",
    keywords: ["turkey", "thanksgiving", "holiday dinner", "holiday meal"],
    types: ["White", "Red", "Sparkling"],
    grapes: ["Chardonnay", "Viognier", "Pinot Noir"]
  },
  {
    label: "pasta/tomato",
    keywords: ["pasta", "tomato sauce", "marinara", "spaghetti", "lasagna", "red sauce"],
    types: ["Red"],
    grapes: ["Sangiovese", "Grenache", "Tempranillo"]
  },
  {
    label: "pizza",
    keywords: ["pizza"],
    types: ["Red", "Sparkling"],
    grapes: ["Sangiovese", "Grenache"]
  },
  {
    label: "spicy food",
    keywords: ["spicy food", "spicy", "curry", "thai food", "indian food", "szechuan", "kimchi"],
    types: ["White", "Sparkling", "Rose"],
    grapes: ["Riesling", "Chenin Blanc"]
  },
  {
    label: "bbq/grilled",
    keywords: ["bbq", "barbecue", "grilled", "burgers", "ribs", "brisket"],
    types: ["Red"],
    grapes: ["Zinfandel", "Syrah", "Malbec", "Grenache"]
  },
  {
    label: "cheese board",
    keywords: ["cheese", "cheese board", "charcuterie", "charcuterie board"],
    types: ["Red", "White"],
    grapes: ["Cabernet Sauvignon", "Sangiovese", "Chardonnay"]
  },
  {
    label: "chocolate/dessert",
    keywords: ["chocolate", "dessert", "cake", "sweets", "tiramisu"],
    types: ["Dessert", "Red"],
    grapes: ["Zinfandel", "Semillon", "Muscat"]
  },
  {
    label: "salad/veggie",
    keywords: ["salad", "vegetables", "vegetarian", "veggie", "veggies"],
    types: ["White", "Rose"],
    grapes: ["Sauvignon Blanc", "Albarino"]
  },
  {
    label: "mushroom/earthy",
    keywords: ["mushroom", "mushrooms", "earthy", "truffle"],
    types: ["Red"],
    grapes: ["Pinot Noir", "Nebbiolo"]
  },

  // --- Occasions / mood ---
  {
    label: "celebration",
    keywords: ["celebration", "toast", "anniversary", "special occasion", "engagement", "wedding"],
    types: ["Sparkling"],
    ratingMin: 92
  },
  {
    label: "summer/patio",
    keywords: ["picnic", "summer", "patio", "poolside", "warm weather", "backyard"],
    types: ["White", "Rose", "Sparkling"]
  },
  {
    label: "cozy/winter",
    keywords: ["cozy", "winter", "fireplace", "cold night", "cold weather"],
    types: ["Red"],
    grapes: ["Cabernet Sauvignon", "Syrah", "Malbec"]
  },
  {
    label: "date night",
    keywords: ["date night", "romantic"],
    types: ["Sparkling", "Rose", "Red"]
  },
  {
    label: "brunch",
    keywords: ["brunch", "mimosa", "mimosas"],
    types: ["Sparkling"]
  },
  {
    label: "gift/splurge",
    keywords: ["gift", "impress", "fancy", "splurge", "treat myself"],
    ratingMin: 93
  },
  {
    label: "budget-friendly",
    keywords: ["budget", "cheap", "inexpensive", "affordable", "everyday", "casual weeknight", "bargain"],
    priceMax: 15
  },

  // --- Style / body descriptors ---
  {
    label: "light/crisp",
    keywords: ["crisp", "light bodied", "light-bodied", "refreshing"],
    types: ["White", "Sparkling"]
  },
  {
    label: "bold/full-bodied",
    keywords: ["bold", "full bodied", "full-bodied", "robust", "big red"],
    types: ["Red"],
    grapes: ["Cabernet Sauvignon", "Syrah", "Malbec", "Zinfandel", "Nebbiolo"]
  },
  {
    label: "smooth/mellow",
    keywords: ["smooth", "mellow", "easy drinking", "easy-drinking"],
    grapes: ["Merlot", "Pinot Noir", "Tempranillo"]
  },
  {
    label: "sweet",
    keywords: ["sweet wine", "dessert wine"],
    types: ["Dessert"]
  },
  {
    label: "dry",
    keywords: ["dry wine", "bone dry", "dry"],
    excludeTypes: ["Dessert"]
  },
  {
    label: "fruity",
    keywords: ["fruity", "fruit forward", "fruit-forward"],
    grapes: ["Zinfandel", "Grenache", "Malbec", "Riesling"]
  }
];