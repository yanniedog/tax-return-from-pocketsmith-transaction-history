import express from "express";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 3000);
const CACHE_PATH = path.join(__dirname, "merchant-intel-cache.json");
const LOOKUP_HEADERS = {
  "user-agent": "Mozilla/5.0 (compatible; PocketSmithTaxPrep/1.0; +https://localhost)",
  accept: "text/html,application/xhtml+xml"
};

const CATEGORY_RULES = [
  {
    category: "health_employer",
    businessType: "Health service / hospital employer",
    keywords: ["nsw health", "health service", "hospital", "local health district", "st vincent", "healthscope"]
  },
  {
    category: "staffing_agency",
    businessType: "Staffing / recruitment agency",
    keywords: ["medrecruit", "recruit", "staffing", "locum", "medical recruitment", "employment agency"]
  },
  {
    category: "government_employer",
    businessType: "Government body / public authority",
    keywords: ["department of", "australian government", "state of", "city council", "government", "services australia"]
  },
  {
    category: "tax_accounting",
    businessType: "Tax or accounting service",
    keywords: ["tax agent", "accountant", "bookkeeping", "tax return", "h&r block", "etax"]
  },
  {
    category: "office_supplies",
    businessType: "Office supplies / stationery retail",
    keywords: ["officeworks", "office supplies", "stationery", "office furniture", "business supplies"]
  },
  {
    category: "software_technology",
    businessType: "Software / technology services",
    keywords: ["software", "google", "microsoft", "adobe", "github", "domain", "cloud", "hosting", "saas"]
  },
  {
    category: "telecom_internet",
    businessType: "Telecommunications / internet provider",
    keywords: ["telstra", "optus", "vodafone", "internet", "nbn", "telecommunications", "mobile"]
  },
  {
    category: "education_training",
    businessType: "Education / training provider",
    keywords: ["training", "course", "academy", "university", "udemy", "coursera", "education"]
  },
  {
    category: "professional_services",
    businessType: "Professional body / insurance / advisory",
    keywords: ["mda", "professional indemnity", "insurance", "association", "college", "ranzco", "ahpra"]
  },
  {
    category: "banking_financial",
    businessType: "Bank / financial services",
    keywords: ["bank", "amex", "american express", "credit card", "finance", "financial", "anz", "nab", "ing"]
  },
  {
    category: "food_beverage",
    businessType: "Food and beverage",
    keywords: ["pizza", "restaurant", "uber eats", "cafe", "food delivery", "dining", "bar", "pub"]
  },
  {
    category: "grocery_retail",
    businessType: "Grocery retail",
    keywords: ["woolworths", "coles", "aldi", "grocery", "supermarket"]
  },
  {
    category: "general_retail",
    businessType: "General consumer retail",
    keywords: ["retail", "shop", "store", "kmart", "target", "amazon"]
  },
  {
    category: "entertainment_media",
    businessType: "Entertainment / media subscription",
    keywords: ["youtube", "netflix", "spotify", "streaming", "playstore", "media"]
  },
  {
    category: "personal_travel",
    businessType: "Travel / transport service",
    keywords: ["transport", "uber", "didi", "taxi", "airlines", "flight", "opal", "transport for nsw"]
  },
  {
    category: "housing_private",
    businessType: "Housing or private finance",
    keywords: ["mortgage", "loan", "real estate", "property", "rent", "home loan", "offset"]
  },
  {
    category: "investment_finance",
    businessType: "Investment / brokerage platform",
    keywords: ["spaceship", "raiz", "broker", "invest", "capital", "wealth", "trading"]
  }
];

const app = express();
app.use(express.json({ limit: "2mb" }));
app.use(express.static(__dirname, { extensions: ["html"] }));

const merchantCache = await loadCache();
let saveTimer = null;

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, cacheEntries: Object.keys(merchantCache).length });
});

app.post("/api/enrich-merchants", async (req, res) => {
  const merchants = Array.isArray(req.body?.merchants) ? req.body.merchants : [];
  const forceRefresh = Boolean(req.body?.forceRefresh);

  if (!merchants.length) {
    res.status(400).json({ error: "No merchants supplied." });
    return;
  }

  if (merchants.length > 250) {
    res.status(400).json({ error: "Batch too large. Limit 250 merchants per request." });
    return;
  }

  const unique = dedupeMerchants(merchants);
  const items = await runWithConcurrency(unique, 3, async (merchant) => {
    try {
      return await enrichMerchant(merchant, forceRefresh);
    } catch (error) {
      return {
        lookupKey: merchant.lookupKey || deriveMerchantLookupKey(merchant.merchant || ""),
        merchantRaw: merchant.merchant || "",
        merchantLookupName: deriveLookupName(merchant.merchant || ""),
        businessType: "Unknown - enrichment failed",
        businessCategory: "unknown",
        classificationConfidence: "low",
        classificationReason: error.message || "Lookup failed.",
        abn: "",
        abnName: "",
        abnEntityType: "",
        abnStatus: "",
        mainPlaceOfBusiness: "",
        sourceUrls: [],
        updatedAt: new Date().toISOString()
      };
    }
  });

  res.json({ items, total: items.length });
});

app.listen(PORT, () => {
  console.log(`PocketSmith tax prep app running at http://localhost:${PORT}`);
});

function dedupeMerchants(merchants) {
  const map = new Map();

  for (const merchant of merchants) {
    const merchantRaw = String(merchant?.merchant || "").trim();
    const lookupKey = String(merchant?.lookupKey || deriveMerchantLookupKey(merchantRaw)).trim();
    if (!lookupKey) {
      continue;
    }

    if (!map.has(lookupKey)) {
      map.set(lookupKey, {
        lookupKey,
        merchant: merchantRaw || lookupKey
      });
    }
  }

  return Array.from(map.values());
}

async function enrichMerchant(merchant, forceRefresh) {
  const lookupKey = merchant.lookupKey || deriveMerchantLookupKey(merchant.merchant || "");
  const merchantRaw = String(merchant.merchant || "").trim() || lookupKey;
  const merchantLookupName = deriveLookupName(merchantRaw);

  if (!lookupKey) {
    throw new Error("No merchant lookup key available.");
  }

  if (!forceRefresh && merchantCache[lookupKey]) {
    return merchantCache[lookupKey];
  }

  const sourceUrls = [];

  let abrResult = null;
  const abrQueries = buildAbrQueries(merchantLookupName);
  let bestAbr = null;
  for (const abrQuery of abrQueries) {
    try {
      const candidate = await lookupAbrByName(abrQuery, merchantLookupName);
      if (!candidate) {
        continue;
      }
      if (!bestAbr || candidate.matchScore > bestAbr.matchScore) {
        bestAbr = candidate;
      }
    } catch (_error) {
      continue;
    }
  }
  abrResult = bestAbr;
  if (abrResult?.searchUrl) {
    sourceUrls.push(abrResult.searchUrl);
  }

  let abrDetails = null;
  if (abrResult?.abn) {
    try {
      abrDetails = await fetchAbrDetails(abrResult.abn);
      if (abrDetails?.detailsUrl) {
        sourceUrls.push(abrDetails.detailsUrl);
      }
    } catch (_error) {
      abrDetails = null;
    }
  }

  let search = null;
  try {
    search = await searchWebBusinessInfo(abrResult?.name || merchantLookupName);
    if (search?.queryUrl) {
      sourceUrls.push(search.queryUrl);
    }
    for (const result of search?.results || []) {
      if (result.url) {
        sourceUrls.push(result.url);
      }
    }
  } catch (_error) {
    search = null;
  }

  const classification = classifyBusinessCategory({
    merchantRaw,
    merchantLookupName,
    abrResult,
    abrDetails,
    search
  });

  const enriched = {
    lookupKey,
    merchantRaw,
    merchantLookupName,
    businessType: classification.businessType,
    businessCategory: classification.businessCategory,
    classificationConfidence: classification.confidence,
    classificationReason: classification.reason,
    abn: abrResult?.abn || "",
    abnName: abrDetails?.entityName || abrResult?.name || "",
    abnEntityType: abrDetails?.entityType || "",
    abnStatus: abrDetails?.abnStatus || abrResult?.status || "",
    mainPlaceOfBusiness: abrDetails?.mainBusinessLocation || abrResult?.location || "",
    sourceUrls: uniqueCompact(sourceUrls).slice(0, 6),
    updatedAt: new Date().toISOString()
  };

  merchantCache[lookupKey] = enriched;
  scheduleCacheSave();
  return enriched;
}

async function lookupAbrByName(name, merchantHint) {
  const query = encodeURIComponent(name);
  const searchUrl = `https://abr.business.gov.au/Search/ResultsActive?SearchText=${query}`;
  const html = await fetchText(searchUrl);

  const rawMatches = Array.from(html.matchAll(/id="Results_NameItems_\d+__Compressed"[^>]*value="([^"]+)"/g));
  const rows = rawMatches
    .map((match) => decodeHtml(match[1]))
    .map(parseAbrCompressedRow)
    .filter(Boolean);

  if (!rows.length) {
    return null;
  }

  const cleanedName = normalizeText(name);
  const hintToken = normalizeText(merchantHint || "").split(" ").find(Boolean) || "";
  const ranked = rows
    .map((row) => {
      const similarity = tokenSimilarity(cleanedName, normalizeText(row.name));
      const includes = normalizeText(row.name).includes(cleanedName) ? 1 : 0;
      const hintBoost = hintToken && normalizeText(row.name).includes(hintToken) ? 1 : 0;
      const nameTypeBoost = row.nameType === "Entity Name" ? 1 : 0;
      return {
        ...row,
        rank: row.relevance + similarity * 35 + includes * 10 + hintBoost * 18 + nameTypeBoost * 5
      };
    })
    .sort((a, b) => b.rank - a.rank);

  const best = ranked[0];
  const confidence = ranked.length > 1 && best.rank - ranked[1].rank < 4 ? "medium" : "high";

  return {
    abn: best.abn,
    name: best.name,
    status: best.status,
    nameType: best.nameType,
    location: best.location,
    state: best.state,
    postcode: best.postcode,
    relevance: best.relevance,
    confidence,
    matchScore: best.rank,
    searchUrl
  };
}

function parseAbrCompressedRow(raw) {
  const parts = raw.split(",");
  if (parts.length < 14) {
    return null;
  }

  const abn = parts[0].replace(/\s+/g, "").trim();
  const status = collapseWhitespace(parts[3]);
  const name = collapseWhitespace(parts[5]);
  const nameType = collapseWhitespace(parts[8]);
  const location = collapseWhitespace(parts[10]);
  const state = collapseWhitespace(parts[11]);
  const postcode = collapseWhitespace(parts[12]);
  const relevance = Number(collapseWhitespace(parts[13])) || 0;

  if (!abn || !name) {
    return null;
  }

  return {
    abn,
    status,
    name,
    nameType,
    location,
    state,
    postcode,
    relevance
  };
}

async function fetchAbrDetails(abn) {
  const detailsUrl = `https://abr.business.gov.au/ABN/View?abn=${encodeURIComponent(abn)}`;
  const html = await fetchText(detailsUrl);

  return {
    entityName: extractTableValue(html, "Entity name"),
    abnStatus: extractTableValue(html, "ABN status"),
    entityType: extractTableValue(html, "Entity type"),
    gstStatus: extractTableValue(html, "Goods & Services Tax (GST)"),
    mainBusinessLocation: extractTableValue(html, "Main business location"),
    detailsUrl
  };
}

function extractTableValue(html, label) {
  const escapedLabel = escapeRegExp(label);
  const pattern = new RegExp(`<th[^>]*>\\s*${escapedLabel}:?\\s*<\\/th>\\s*<td[^>]*>([\\s\\S]*?)<\\/td>`, "i");
  const match = pattern.exec(html);
  if (!match) {
    return "";
  }
  return collapseWhitespace(stripHtml(decodeHtml(match[1])));
}

async function searchWebBusinessInfo(name) {
  const query = `${name} Australia business`;
  const queryUrl = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`;
  const html = await fetchText(queryUrl);

  const titleMatches = Array.from(html.matchAll(/class="result__a"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g));
  const snippetMatches = Array.from(html.matchAll(/class="result__snippet"[^>]*>([\s\S]*?)<\/(?:a|div)>/g));

  const results = [];
  const count = Math.min(8, titleMatches.length);
  for (let i = 0; i < count; i += 1) {
    const title = collapseWhitespace(stripHtml(decodeHtml(titleMatches[i][2])));
    const url = unwrapDuckDuckGoUrl(decodeHtml(titleMatches[i][1]));
    const snippet = snippetMatches[i] ? collapseWhitespace(stripHtml(decodeHtml(snippetMatches[i][1]))) : "";

    if (!title) {
      continue;
    }

    results.push({ title, url, snippet });
  }

  return {
    query,
    queryUrl,
    results
  };
}

function unwrapDuckDuckGoUrl(rawUrl) {
  if (!rawUrl) {
    return "";
  }

  let value = rawUrl;
  if (value.startsWith("//")) {
    value = `https:${value}`;
  }

  try {
    const parsed = new URL(value);
    if (parsed.hostname.includes("duckduckgo.com") && parsed.searchParams.has("uddg")) {
      return decodeURIComponent(parsed.searchParams.get("uddg"));
    }
    return parsed.toString();
  } catch (_error) {
    return rawUrl;
  }
}

function classifyBusinessCategory({ merchantRaw, merchantLookupName, abrResult, abrDetails, search }) {
  const text = normalizeText(
    [
      merchantRaw,
      merchantLookupName,
      abrResult?.name,
      abrDetails?.entityName,
      abrDetails?.entityType,
      ...((search?.results || []).map((item) => `${item.title} ${item.snippet}`))
    ]
      .filter(Boolean)
      .join(" ")
  );

  if (!text) {
    return {
      businessType: "Unknown",
      businessCategory: "unknown",
      confidence: "low",
      reason: "No merchant text available for classification."
    };
  }

  let best = null;

  for (const rule of CATEGORY_RULES) {
    let score = 0;
    for (const keyword of rule.keywords) {
      if (text.includes(keyword)) {
        score += keyword.includes(" ") ? 2 : 1;
      }
    }

    if (score <= 0) {
      continue;
    }

    const weightedScore = score + (abrResult?.abn ? 1 : 0);
    if (!best || weightedScore > best.score) {
      best = {
        ...rule,
        score: weightedScore
      };
    }
  }

  if (!best) {
    return {
      businessType: abrDetails?.entityType || "Unknown",
      businessCategory: "unknown",
      confidence: abrResult?.abn ? "medium" : "low",
      reason: abrResult?.abn
        ? "ABN found but business activity type was not confidently inferred from search text."
        : "No reliable keyword match from ABR/web sources."
    };
  }

  const confidence = best.score >= 4 ? "high" : best.score >= 2 ? "medium" : "low";
  const reasonParts = [`Matched merchant/profile keywords for ${best.businessType}.`];
  if (abrResult?.abn) {
    reasonParts.push("ABN match found in ABR results.");
  }

  return {
    businessType: best.businessType,
    businessCategory: best.category,
    confidence,
    reason: reasonParts.join(" ")
  };
}

async function fetchText(url) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: LOOKUP_HEADERS,
      redirect: "follow",
      signal: controller.signal
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status} for ${url}`);
    }

    return await response.text();
  } finally {
    clearTimeout(timeout);
  }
}

function buildAbrQueries(name) {
 const normalized = normalizeText(name || "");
 const tokens = normalized.split(" ").filter((token) => token.length > 1);
 if (!tokens.length) {
   return [];
 }

 const legalWords = new Set(["pty", "ltd", "limited", "company", "co", "the", "trustee", "for", "trust", "unit"]);
 const trimmed = tokens.filter((token) => !legalWords.has(token));

 const candidates = [
   tokens.join(" "),
   tokens.slice(0, 4).join(" "),
   tokens.slice(0, 3).join(" "),
   tokens.slice(0, 2).join(" "),
   tokens[0],
   trimmed.join(" "),
   trimmed.slice(0, 3).join(" "),
   trimmed[0]
 ];

 return uniqueCompact(candidates).filter(Boolean);
}

function deriveLookupName(merchantRaw) {
  const raw = normalizeText(merchantRaw || "");
  if (!raw) {
    return "";
  }

  let name = raw;
  name = name.replace(/^paypal\s+/, "");
  name = name.replace(/^google\s+\*?\s*/, "google ");
  name = name.replace(/^transportfornsw/, "transport for nsw");
  name = name.replace(/\b(online|payment|received|thankyou|receipt|debit|credit|card|date|from|to|ref)\b/g, " ");
  name = name.replace(/\b(x{2,}\d*|\d{3,})\b/g, " ");
  name = name.replace(/\s+/g, " ").trim();

  const tokens = name.split(" ").filter((token) => token.length > 1);
  return tokens.slice(0, 7).join(" ");
}

function deriveMerchantLookupKey(merchantRaw) {
  return deriveLookupName(merchantRaw).slice(0, 120);
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[\\/\-_*]+/g, " ")
    .replace(/[^a-z0-9.& ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenSimilarity(left, right) {
  const a = new Set(String(left || "").split(" ").filter((token) => token.length > 2));
  const b = new Set(String(right || "").split(" ").filter((token) => token.length > 2));

  if (!a.size || !b.size) {
    return 0;
  }

  let intersection = 0;
  for (const token of a) {
    if (b.has(token)) {
      intersection += 1;
    }
  }

  const union = new Set([...a, ...b]).size;
  return union ? intersection / union : 0;
}

function stripHtml(value) {
  return String(value || "").replace(/<[^>]*>/g, " ");
}

function decodeHtml(value) {
  const named = {
    "&nbsp;": " ",
    "&amp;": "&",
    "&quot;": '"',
    "&#39;": "'",
    "&#x27;": "'",
    "&lt;": "<",
    "&gt;": ">"
  };

  let text = String(value || "");
  for (const [entity, replacement] of Object.entries(named)) {
    text = text.split(entity).join(replacement);
  }

  text = text.replace(/&#(\d+);/g, (_m, code) => String.fromCharCode(Number(code)));
  text = text.replace(/&#x([0-9a-fA-F]+);/g, (_m, code) => String.fromCharCode(parseInt(code, 16)));

  return text;
}

function collapseWhitespace(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function uniqueCompact(values) {
  return Array.from(new Set(values.filter(Boolean).map((value) => String(value).trim()).filter(Boolean)));
}

function escapeRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function runWithConcurrency(items, concurrency, worker) {
  const limit = Math.max(1, Math.min(concurrency, items.length || 1));
  const output = new Array(items.length);
  let nextIndex = 0;

  const runners = Array.from({ length: limit }, async () => {
    while (true) {
      const current = nextIndex;
      if (current >= items.length) {
        break;
      }
      nextIndex += 1;
      output[current] = await worker(items[current], current);
    }
  });

  await Promise.all(runners);
  return output;
}

async function loadCache() {
  try {
    const raw = await fs.readFile(CACHE_PATH, "utf8");
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") {
      return parsed;
    }
    return {};
  } catch (_error) {
    return {};
  }
}

function scheduleCacheSave() {
  if (saveTimer) {
    clearTimeout(saveTimer);
  }

  saveTimer = setTimeout(async () => {
    try {
      await fs.writeFile(CACHE_PATH, JSON.stringify(merchantCache, null, 2));
    } catch (error) {
      console.error("Failed to save merchant cache", error);
    }
  }, 500);
}
