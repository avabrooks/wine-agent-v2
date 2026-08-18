import { TYPE_SYNONYMS, HEADER_SYNONYMS } from "./constants";
import { COUNTRY_SYNONYMS } from "./countries";
import { PAIRING_DICTIONARY } from "./pairings";

export interface ValuesFilter {
  column: string; // actual header text, e.g. "Type"
  values: string[]; // exact cell text values to show
}

export interface CustomFilter {
  column: string;
  criteria1: string;
  criteria2?: string;
  oper?: "And" | "Or";
}

export interface FilterPlan {
  valuesFilters: ValuesFilter[];
  customFilters: CustomFilter[];
  sort?: { column: string; ascending: boolean };
}

export interface ParseResult {
  plan: FilterPlan;
  notes: string[];
  // Estimated match count, computed locally against the data we read from the
  // sheet. Excel's own post-filter visible-row count is used for the status
  // message instead once the plan is actually applied (see taskpane.ts).
  count: number;
}

interface ColInfo {
  name: string; // actual header text in the sheet
  index: number; // column index within the data body range
}

type ColMap = Record<string, ColInfo>;

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function wordMatch(promptLower: string, value: string): boolean {
  const re = new RegExp("\\b" + escapeRegex(value.toLowerCase()) + "\\b");
  return re.test(promptLower);
}

const NEGATION_CUES = "(?:not|no|except|excluding|other than|aside from|besides)";

// True when a negation cue phrase sits directly in front of this specific
// term — "not from the US", "no French wines", "excluding sparkling",
// "besides Malbec", "non-Italian" — rather than just "does a negation word
// appear anywhere in the prompt". Scoping it to the term itself means a
// prompt that both includes one value and excludes another in the same
// field ("red, not sparkling") resolves each independently.
function isNegated(promptLower: string, term: string): boolean {
  const t = escapeRegex(term.toLowerCase());
  const cueBefore = new RegExp("\\b" + NEGATION_CUES + "\\s+(?:from\\s+|the\\s+|a\\s+|an\\s+)*" + t + "\\b");
  const nonPrefixed = new RegExp("\\bnon-?\\s*" + t + "\\b");
  return cueBefore.test(promptLower) || nonPrefixed.test(promptLower);
}

// Strips punctuation/currency symbols down to letters, digits, spaces and a
// couple of separators, so "Retail Price (USD)" and "retail-price" both
// normalize to something comparable against the synonym lists.
function normalizeHeader(h: any): string {
  return String(h ?? "")
    .toLowerCase()
    .replace(/[()]/g, " ")
    .replace(/[^a-z0-9\s/$-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp: number[][] = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] =
        a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
    }
  }
  return dp[m][n];
}

// Maps a sheet's actual header row onto our canonical fields (name, type,
// price, ...). Tries three passes, each looser than the last, so headers
// that don't exactly match HEADER_SYNONYMS still get recognized:
//   1. exact match against a synonym
//   2. substring match either direction ("Retail Price (USD)" contains "price")
//   3. fuzzy spelling match (edit distance <= 2) for near-miss headers/typos
// Once a column is claimed by one canonical field it can't also be claimed
// by another, so ambiguous headers don't get double-booked.
function detectColumns(headers: any[]): ColMap {
  const map: ColMap = {};
  const normalized = headers.map((h) => normalizeHeader(h));
  const claimed = new Set<number>();

  const claim = (canon: string, idx: number) => {
    map[canon] = { name: String(headers[idx]), index: idx };
    claimed.add(idx);
  };

  for (const canon in HEADER_SYNONYMS) {
    const syns = HEADER_SYNONYMS[canon];

    let idx = normalized.findIndex((h, i) => !claimed.has(i) && syns.includes(h));

    if (idx === -1) {
      idx = normalized.findIndex(
        (h, i) => !claimed.has(i) && h.length > 0 && syns.some((s) => h.includes(s) || s.includes(h))
      );
    }

    if (idx === -1) {
      idx = normalized.findIndex((h, i) => {
        if (claimed.has(i) || h.length < 3) return false;
        return syns.some((s) => Math.abs(h.length - s.length) <= 2 && levenshtein(h, s) <= 2);
      });
    }

    if (idx !== -1) claim(canon, idx);
  }
  return map;
}

// Pulls a plain number out of messy cell text like "$45.00", "1,200",
// "92 pts", or "4.5 stars" — used only for the local match-count estimate
// and notes, since the actual Excel-side filter/sort compares real cell
// values (see columnLooksTextual below for when that can go wrong).
function parseNumeric(v: any): number | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  const s = String(v).replace(/[^0-9.-]/g, "");
  if (!s || s === "-" || s === ".") return null;
  const n = parseFloat(s);
  return Number.isNaN(n) ? null : n;
}

// True when a column's values are stored as text rather than numbers (e.g. a
// price column pasted in as "$45.00" strings). Excel's native Table filter
// and sort compare real cell values, so a numeric filter/sort we build here
// may not apply correctly against that column — we can't silently fix it
// without rewriting the user's cells, so parseAndApply surfaces a note
// instead when this is detected.
function columnLooksTextual(rows: any[][], col: ColInfo): boolean {
  const sample = rows
    .slice(0, 30)
    .map((r) => r[col.index])
    .filter((v) => v !== "" && v !== null && v !== undefined);
  if (!sample.length) return false;
  return sample.every((v) => typeof v !== "number");
}

function uniqueValues(rows: any[][], col: ColInfo | undefined): string[] {
  if (!col) return [];
  const set = new Set<string>();
  rows.forEach((r) => {
    const v = r[col.index];
    if (v !== undefined && v !== null && v !== "") set.add(String(v));
  });
  return Array.from(set).sort((a, b) => b.length - a.length);
}

const NOT_POINTS = "(?!\\s*(?:points?|pts?|\\/100|out of))";
const NOT_RATING_WORD =
  "(?<!\\brated\\s)(?<!\\brating\\s)(?<!\\bscored\\s)(?<!\\bscore\\s)";

export function parseAndApply(headers: any[], originalRows: any[][], prompt: string): ParseResult {
  const colMap = detectColumns(headers);
  const p = " " + prompt.toLowerCase().trim() + " ";
  const notes: string[] = [];

  const plan: FilterPlan = { valuesFilters: [], customFilters: [] };

  // rows is only used locally to estimate a match count / build notes; the
  // actual filtering in Excel happens via the Table filter API in the plan.
  let rows = originalRows.slice();

  let hadDirectType = false;
  let hadDirectGrape = false;
  let hadExplicitPrice = false;
  let hadExplicitRating = false;

  // ---- Type / color ----
  if (colMap.type) {
    const matchedCanon: string[] = [];
    const excludedCanon: string[] = [];
    for (const canon in TYPE_SYNONYMS) {
      const syns = TYPE_SYNONYMS[canon];
      if (syns.some((syn) => wordMatch(p, syn))) {
        if (syns.some((syn) => isNegated(p, syn))) excludedCanon.push(canon);
        else matchedCanon.push(canon);
      }
    }
    if (matchedCanon.length || excludedCanon.length) {
      const dataTypes = uniqueValues(originalRows, colMap.type);
      const canonToActual: Record<string, string> = {};
      dataTypes.forEach((dt) => {
        const dtl = dt.toLowerCase();
        for (const canon in TYPE_SYNONYMS) {
          if (TYPE_SYNONYMS[canon].some((s) => dtl.includes(s))) canonToActual[canon] = dt;
        }
      });
      const actualMatches = [...new Set(matchedCanon.map((c) => canonToActual[c]).filter(Boolean))];
      const actualExcluded = [...new Set(excludedCanon.map((c) => canonToActual[c]).filter(Boolean))];
      const base = actualMatches.length ? actualMatches : dataTypes;
      const keep = base.filter((t) => !actualExcluded.includes(t));

      if (keep.length) {
        rows = rows.filter((r) => keep.includes(String(r[colMap.type.index])));
        plan.valuesFilters.push({ column: colMap.type.name, values: keep });
        if (actualMatches.length && actualExcluded.length) {
          notes.push(`Type: ${actualMatches.join(" or ")} (not ${actualExcluded.join(" or ")})`);
        } else if (actualExcluded.length) {
          notes.push(`Type: not ${actualExcluded.join(" or ")}`);
        } else {
          notes.push(`Type: ${actualMatches.join(" or ")}`);
        }
        hadDirectType = true;
      }
    }
  }

  // ---- Country ----
  if (colMap.country) {
    const dataCountries = uniqueValues(originalRows, colMap.country);
    const matchedCountryNames: string[] = [];
    const excludedCountryNames: string[] = [];
    for (const canonCountry in COUNTRY_SYNONYMS) {
      const syns = COUNTRY_SYNONYMS[canonCountry];
      if (syns.some((syn) => wordMatch(p, syn))) {
        const actual = dataCountries.find((c) => c.toLowerCase() === canonCountry.toLowerCase());
        if (actual) {
          if (syns.some((syn) => isNegated(p, syn))) excludedCountryNames.push(actual);
          else matchedCountryNames.push(actual);
        }
      }
    }
    if (!matchedCountryNames.length && !excludedCountryNames.length) {
      dataCountries.forEach((c) => {
        if (wordMatch(p, c)) {
          if (isNegated(p, c)) excludedCountryNames.push(c);
          else matchedCountryNames.push(c);
        }
      });
    }
    const countryBase = matchedCountryNames.length ? matchedCountryNames : dataCountries;
    const countryKeep = countryBase.filter((c) => !excludedCountryNames.includes(c));

    if (countryKeep.length && (matchedCountryNames.length || excludedCountryNames.length)) {
      rows = rows.filter((r) => countryKeep.includes(String(r[colMap.country.index])));
      plan.valuesFilters.push({ column: colMap.country.name, values: countryKeep });
      if (matchedCountryNames.length && excludedCountryNames.length) {
        notes.push(`Country: ${matchedCountryNames.join(" or ")} (not ${excludedCountryNames.join(" or ")})`);
      } else if (excludedCountryNames.length) {
        notes.push(`Country: not ${excludedCountryNames.join(" or ")}`);
      } else {
        notes.push(`Country: ${matchedCountryNames.join(" or ")}`);
      }
    }
  }

  // ---- Region ----
  if (colMap.region) {
    const regions = uniqueValues(originalRows, colMap.region);
    const matched = regions.filter((rgn) => wordMatch(p, rgn) && !isNegated(p, rgn));
    const excluded = regions.filter((rgn) => wordMatch(p, rgn) && isNegated(p, rgn));
    const base = matched.length ? matched : regions;
    const keep = base.filter((rgn) => !excluded.includes(rgn));

    if (keep.length && (matched.length || excluded.length)) {
      rows = rows.filter((r) => keep.includes(String(r[colMap.region.index])));
      plan.valuesFilters.push({ column: colMap.region.name, values: keep });
      if (matched.length && excluded.length) {
        notes.push(`Region: ${matched.join(" or ")} (not ${excluded.join(" or ")})`);
      } else if (excluded.length) {
        notes.push(`Region: not ${excluded.join(" or ")}`);
      } else {
        notes.push(`Region: ${matched.join(" or ")}`);
      }
    }
  }

  // ---- Grape ----
  if (colMap.grape) {
    const grapes = uniqueValues(originalRows, colMap.grape);
    const grapeHit = (g: string) => wordMatch(p, g) || wordMatch(p, g.split(" ")[0]);
    const grapeNegated = (g: string) => isNegated(p, g) || isNegated(p, g.split(" ")[0]);
    const matched = grapes.filter((g) => grapeHit(g) && !grapeNegated(g));
    const excluded = grapes.filter((g) => grapeHit(g) && grapeNegated(g));
    const base = matched.length ? matched : grapes;
    const keep = base.filter((g) => !excluded.includes(g));

    if (keep.length && (matched.length || excluded.length)) {
      rows = rows.filter((r) => keep.includes(String(r[colMap.grape.index])));
      plan.valuesFilters.push({ column: colMap.grape.name, values: keep });
      if (matched.length && excluded.length) {
        notes.push(`Grape: ${matched.join(" or ")} (not ${excluded.join(" or ")})`);
      } else if (excluded.length) {
        notes.push(`Grape: not ${excluded.join(" or ")}`);
      } else {
        notes.push(`Grape: ${matched.join(" or ")}`);
      }
      hadDirectGrape = true;
    }
  }

  // ---- Price ----
  if (colMap.price) {
    let m: RegExpMatchArray | null;
    if (
      (m = p.match(
        new RegExp("between\\s*\\$?(\\d+(?:\\.\\d+)?)\\b\\s*(?:and|-|to)\\s*\\$?(\\d+(?:\\.\\d+)?)\\b" + NOT_POINTS)
      ))
    ) {
      const lo = parseFloat(m[1]);
      const hi = parseFloat(m[2]);
      rows = rows.filter((r) => {
        const v = parseNumeric(r[colMap.price.index]);
        return v !== null && v >= lo && v <= hi;
      });
      plan.customFilters.push({ column: colMap.price.name, criteria1: `>=${lo}`, criteria2: `<=${hi}`, oper: "And" });
      notes.push(`Price: $${lo}–$${hi}`);
      hadExplicitPrice = true;
    } else if (
      (m = p.match(
        new RegExp(NOT_RATING_WORD + "(?:under|below|less than|cheaper than|up to)\\s*\\$?(\\d+(?:\\.\\d+)?)\\b" + NOT_POINTS)
      ))
    ) {
      const max = parseFloat(m[1]);
      rows = rows.filter((r) => {
        const v = parseNumeric(r[colMap.price.index]);
        return v !== null && v <= max;
      });
      plan.customFilters.push({ column: colMap.price.name, criteria1: `<=${max}` });
      notes.push(`Price: ≤ $${max}`);
      hadExplicitPrice = true;
    } else if (
      (m = p.match(
        new RegExp(NOT_RATING_WORD + "(?:over|above|more than|pricier than|at least)\\s*\\$?(\\d+(?:\\.\\d+)?)\\b" + NOT_POINTS)
      ))
    ) {
      const min = parseFloat(m[1]);
      rows = rows.filter((r) => {
        const v = parseNumeric(r[colMap.price.index]);
        return v !== null && v >= min;
      });
      plan.customFilters.push({ column: colMap.price.name, criteria1: `>=${min}` });
      notes.push(`Price: ≥ $${min}`);
      hadExplicitPrice = true;
    }
    if (hadExplicitPrice && columnLooksTextual(originalRows, colMap.price)) {
      notes.push(
        `Heads up: "${colMap.price.name}" looks like it's stored as text (e.g. "$45.00") rather than numbers, so Excel's filter may not apply correctly — try formatting that column as Number first.`
      );
    }
  }

  // ---- Rating ----
  if (colMap.rating) {
    let m: RegExpMatchArray | null;
    if (
      (m = p.match(
        /(?:rated|rating|scored?|score)\s*(?:above|over|at least)\s*(\d+)|(?:over|above|more than|at least)\s*(\d+)\s*(?:points?|pts?)|(\d+)\s*\+\s*(?:points?|pts?)/
      ))
    ) {
      const min = parseInt(m[1] || m[2] || m[3], 10);
      rows = rows.filter((r) => {
        const v = parseNumeric(r[colMap.rating.index]);
        return v !== null && v >= min;
      });
      plan.customFilters.push({ column: colMap.rating.name, criteria1: `>=${min}` });
      notes.push(`Rating: ≥ ${min}`);
      hadExplicitRating = true;
    } else if (
      (m = p.match(/(?:rated|rating|scored?|score)\s*(?:below|under|less than)\s*(\d+)|(?:under|below|less than)\s*(\d+)\s*(?:points?|pts?)/))
    ) {
      const max = parseInt(m[1] || m[2], 10);
      rows = rows.filter((r) => {
        const v = parseNumeric(r[colMap.rating.index]);
        return v !== null && v <= max;
      });
      plan.customFilters.push({ column: colMap.rating.name, criteria1: `<=${max}` });
      notes.push(`Rating: ≤ ${max}`);
      hadExplicitRating = true;
    }
    if (hadExplicitRating && columnLooksTextual(originalRows, colMap.rating)) {
      notes.push(
        `Heads up: "${colMap.rating.name}" looks like it's stored as text (e.g. "92 pts") rather than numbers, so Excel's filter may not apply correctly — try formatting that column as Number first.`
      );
    }
  }

  // ---- Food pairing / occasion / style ----
  {
    const hits = PAIRING_DICTIONARY.filter((entry) => entry.keywords.some((k) => wordMatch(p, k)));
    if (hits.length) {
      const candidateTypes = new Set<string>();
      const candidateGrapes = new Set<string>();
      const excludeTypes = new Set<string>();
      let priceMaxHint: number | null = null;
      let ratingMinHint: number | null = null;
      const labels: string[] = [];

      hits.forEach((h) => {
        (h.types || []).forEach((t) => candidateTypes.add(t.toLowerCase()));
        (h.grapes || []).forEach((g) => candidateGrapes.add(g.toLowerCase()));
        (h.excludeTypes || []).forEach((t) => excludeTypes.add(t.toLowerCase()));
        if (h.priceMax !== undefined) priceMaxHint = priceMaxHint === null ? h.priceMax : Math.min(priceMaxHint, h.priceMax);
        if (h.ratingMin !== undefined) ratingMinHint = ratingMinHint === null ? h.ratingMin : Math.max(ratingMinHint, h.ratingMin);
        labels.push(h.label);
      });

      if (!hadDirectType && !hadDirectGrape && (candidateTypes.size || candidateGrapes.size) && (colMap.type || colMap.grape)) {
        const actualTypes = colMap.type ? uniqueValues(originalRows, colMap.type).filter((t) => candidateTypes.has(t.toLowerCase())) : [];
        const actualGrapes = colMap.grape ? uniqueValues(originalRows, colMap.grape).filter((g) => candidateGrapes.has(g.toLowerCase())) : [];
        if (actualTypes.length) {
          rows = rows.filter((r) => actualTypes.includes(String(r[colMap.type.index])) || (actualGrapes.length && actualGrapes.includes(String(r[colMap.grape.index]))));
          plan.valuesFilters.push({ column: colMap.type.name, values: actualTypes });
        } else if (actualGrapes.length) {
          rows = rows.filter((r) => actualGrapes.includes(String(r[colMap.grape.index])));
          plan.valuesFilters.push({ column: colMap.grape.name, values: actualGrapes });
        }
      }
      if (excludeTypes.size && colMap.type) {
        const actualExclude = uniqueValues(originalRows, colMap.type).filter((t) => excludeTypes.has(t.toLowerCase()));
        if (actualExclude.length) {
          rows = rows.filter((r) => !actualExclude.includes(String(r[colMap.type.index])));
          // Represent "not these types" as a values filter over everything else.
          const allTypes = uniqueValues(originalRows, colMap.type);
          const keep = allTypes.filter((t) => !actualExclude.includes(t));
          plan.valuesFilters.push({ column: colMap.type.name, values: keep });
        }
      }
      if (priceMaxHint !== null && !hadExplicitPrice && colMap.price) {
        rows = rows.filter((r) => {
          const v = parseNumeric(r[colMap.price.index]);
          return v !== null && v <= (priceMaxHint as number);
        });
        plan.customFilters.push({ column: colMap.price.name, criteria1: `<=${priceMaxHint}` });
      }
      if (ratingMinHint !== null && !hadExplicitRating && colMap.rating) {
        rows = rows.filter((r) => {
          const v = parseNumeric(r[colMap.rating.index]);
          return v !== null && v >= (ratingMinHint as number);
        });
        plan.customFilters.push({ column: colMap.rating.name, criteria1: `>=${ratingMinHint}` });
      }
      if (labels.length) notes.push(`Matched to: ${labels.join(", ")}`);
    }
  }

  // ---- Vintage ----
  if (colMap.vintage) {
    let m: RegExpMatchArray | null;
    if ((m = p.match(/(?:after|since|newer than)\s*(\d{4})/))) {
      const min = parseInt(m[1], 10);
      rows = rows.filter((r) => {
        const v = parseNumeric(r[colMap.vintage.index]);
        return v !== null && v >= min;
      });
      plan.customFilters.push({ column: colMap.vintage.name, criteria1: `>=${min}` });
      notes.push(`Vintage: ${min} or later`);
    } else if ((m = p.match(/(?:before|older than)\s*(\d{4})/))) {
      const max = parseInt(m[1], 10);
      rows = rows.filter((r) => {
        const v = parseNumeric(r[colMap.vintage.index]);
        return v !== null && v <= max;
      });
      plan.customFilters.push({ column: colMap.vintage.name, criteria1: `<=${max}` });
      notes.push(`Vintage: before ${max}`);
    } else if ((m = p.match(/\b(\d{4})\b/))) {
      const yr = parseInt(m[1], 10);
      if (yr >= 1900 && yr <= 2100) {
        rows = rows.filter((r) => parseNumeric(r[colMap.vintage.index]) === yr);
        plan.customFilters.push({ column: colMap.vintage.name, criteria1: `=${yr}` });
        notes.push(`Vintage: ${yr}`);
      }
    }
  }

  // ---- Sort ----
  let sortCol: ColInfo | null = null;
  let sortDir: "asc" | "desc" | null = null;
  let sortLabel = "";
  const wantsAsc = /\bascending\b|\basc\b/.test(p);
  const wantsDesc = /\bdescending\b|\bdesc\b/.test(p);

  if (colMap.price && /cheapest first|lowest price first|by price.*asc/.test(p)) {
    sortCol = colMap.price; sortDir = "asc"; sortLabel = "Price (low → high)";
  } else if (colMap.price && /most expensive first|priciest first|highest price first/.test(p)) {
    sortCol = colMap.price; sortDir = "desc"; sortLabel = "Price (high → low)";
  } else if (colMap.rating && /(top|best|highest) rated|best first/.test(p)) {
    sortCol = colMap.rating; sortDir = "desc"; sortLabel = "Rating (high → low)";
  } else if (colMap.rating && /lowest rated first|worst first/.test(p)) {
    sortCol = colMap.rating; sortDir = "asc"; sortLabel = "Rating (low → high)";
  } else if (colMap.vintage && /newest first|most recent first/.test(p)) {
    sortCol = colMap.vintage; sortDir = "desc"; sortLabel = "Vintage (newest first)";
  } else if (colMap.vintage && /oldest first/.test(p)) {
    sortCol = colMap.vintage; sortDir = "asc"; sortLabel = "Vintage (oldest first)";
  } else if (colMap.name && /alphabetical|by name/.test(p)) {
    sortCol = colMap.name; sortDir = "asc"; sortLabel = "Name (A→Z)";
  } else if (colMap.price && /sort(?:ed)? by price|order(?:ed)? by price/.test(p)) {
    sortCol = colMap.price; sortDir = wantsDesc ? "desc" : "asc";
    sortLabel = "Price (" + (sortDir === "asc" ? "low → high" : "high → low") + ")";
  } else if (colMap.rating && /sort(?:ed)? by rating|order(?:ed)? by rating/.test(p)) {
    sortCol = colMap.rating; sortDir = wantsAsc ? "asc" : "desc";
    sortLabel = "Rating (" + (sortDir === "asc" ? "low → high" : "high → low") + ")";
  } else if (colMap.vintage && /sort(?:ed)? by vintage|order(?:ed)? by vintage|sort(?:ed)? by year/.test(p)) {
    sortCol = colMap.vintage; sortDir = wantsAsc ? "asc" : "desc";
    sortLabel = "Vintage (" + (sortDir === "asc" ? "oldest first" : "newest first") + ")";
  }

  if (sortCol && sortDir) {
    const field = sortCol.index;
    const dir = sortDir;
    const numericSort = sortCol === colMap.price || sortCol === colMap.rating || sortCol === colMap.vintage;
    rows.sort((a, b) => {
      const av = numericSort ? parseNumeric(a[field]) ?? a[field] : a[field];
      const bv = numericSort ? parseNumeric(b[field]) ?? b[field] : b[field];
      if (av < bv) return dir === "asc" ? -1 : 1;
      if (av > bv) return dir === "asc" ? 1 : -1;
      return 0;
    });
    plan.sort = { column: sortCol.name, ascending: sortDir === "asc" };
    notes.push(`Sorted by: ${sortLabel}`);
    if (numericSort && columnLooksTextual(originalRows, sortCol)) {
      notes.push(
        `Heads up: "${sortCol.name}" looks like it's stored as text, so Excel's native sort may not put it in true numeric order — try formatting that column as Number first.`
      );
    }
  }

  if (!notes.length) {
    notes.push("Didn't recognize any filters or sort in that prompt — showing everything.");
  }

  return { plan, notes, count: rows.length };
}