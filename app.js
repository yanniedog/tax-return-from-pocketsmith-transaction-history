
"use strict";

const DAY_MS = 24 * 60 * 60 * 1000;
const CURRENCY = "AUD";
const ENRICH_BATCH_SIZE = 20;

const TRANSFER_KEYWORDS_PRIMARY = [
  "transfer",
  "a2a",
  "osko",
  "funds transfer",
  "round up",
  "receipt to",
  "receipt from",
  "savings maximiser",
  "orange everyday",
  "internal"
];

const SALARY_KEYWORDS = [
  "salary",
  "wage",
  "payroll",
  "employment",
  "nsw health",
  "medrecruit",
  "locum"
];

const INTEREST_KEYWORDS = ["interest", "bonus interest"];
const INVESTMENT_INCOME_KEYWORDS = ["dividend", "distribution", "trust distribution"];
const REFUND_KEYWORDS = ["refund", "reversal", "chargeback", "payment returned", "payment received thankyou"];
const TAX_PAYMENT_KEYWORDS = ["tax office", "bpay tax", "income tax", "ato"];
const TAX_AGENT_KEYWORDS = ["tax agent", "accountant", "tax return", "h&r block", "hr block", "etax"];

const DEDUCTION_RULES = [
  {
    atoLabel: "D5",
    category: "Professional fees, registrations and indemnity",
    confidence: "high",
    defaultTreatment: "deduction_likely",
    note: "Professional registration/indemnity expense.",
    keywords: ["mda", "ahpra", "medical defence", "professional indemnity", "registration", "licence", "license", "ranzco", "union fee", "college fee"]
  },
  {
    atoLabel: "D10",
    category: "Cost of managing tax affairs",
    confidence: "high",
    defaultTreatment: "deduction_likely",
    note: "Tax agent or tax affairs management cost.",
    keywords: TAX_AGENT_KEYWORDS
  },
  {
    atoLabel: "D5",
    category: "Work-related tools, software and office supplies",
    confidence: "medium",
    defaultTreatment: "deduction_possible",
    note: "Requires work-use substantiation and apportionment.",
    keywords: [
      "officeworks",
      "microsoft",
      "adobe",
      "github",
      "google storage",
      "zoom",
      "domain",
      "godaddy",
      "software",
      "laptop",
      "monitor",
      "keyboard",
      "mouse",
      "printer",
      "stationery"
    ]
  },
  {
    atoLabel: "D5",
    category: "Home office internet and phone",
    confidence: "medium",
    defaultTreatment: "deduction_possible",
    note: "Use work-use percentage and records.",
    keywords: ["telstra", "optus", "vodafone", "tpg", "internet", "nbn", "mobile plan", "phone bill"]
  },
  {
    atoLabel: "D4",
    category: "Self-education",
    confidence: "medium",
    defaultTreatment: "deduction_possible",
    note: "Confirm direct link to current employment income.",
    keywords: ["udemy", "coursera", "course", "training", "conference", "seminar", "textbook", "exam fee", "study"]
  },
  {
    atoLabel: "D2",
    category: "Work-related travel",
    confidence: "low",
    defaultTreatment: "deduction_possible",
    note: "Commuting is private unless specific conditions apply.",
    keywords: ["transportfornsw", "opal", "uber", "didi", "taxi", "jetstar", "qantas", "virgin", "flight", "parking", "toll"]
  },
  {
    atoLabel: "D9",
    category: "Gifts or donations",
    confidence: "medium",
    defaultTreatment: "deduction_possible",
    note: "Confirm recipient is DGR and donation eligibility.",
    keywords: ["donation", "charity", "red cross", "salvos", "unicef", "lifeline"]
  }
];

const NON_DEDUCTIBLE_RULES = [
  {
    category: "Private meals and entertainment",
    keywords: ["uber eats", "pizza", "restaurant", "dining", "alcohol", "bar", "pub", "cafe", "takeaway"]
  },
  {
    category: "Personal subscriptions and media",
    keywords: ["youtubepremium", "youtube premium", "spotify", "netflix", "google play", "playstore", "apple.com/bill", "strava"]
  },
  {
    category: "Housing, loans and private finance",
    keywords: ["mortgage", "loan", "drawdown", "rent", "home loan", "offset", "repayment"]
  },
  {
    category: "Groceries and household personal spending",
    keywords: ["woolworths", "coles", "aldi", "groceries", "kmart", "target", "ikea", "household"]
  },
  {
    category: "Investment contributions or personal savings",
    keywords: ["spaceship", "raiz", "saving", "savings", "superannuation top up", "investment transfer"]
  },
  {
    category: "Bank fees and penalties (private)",
    keywords: ["overdue payment", "late fee", "interest charges", "charge for overdue"]
  }
];

const CATEGORY_FALLBACK_RULES = [
  {
    match: ["transfers", "savings"],
    treatment: "internal_transfer",
    category: "Likely internal transfer (category fallback)",
    confidence: "low",
    reason: "PocketSmith category fallback used; merchant text was inconclusive."
  },
  {
    match: ["transport", "travel", "computing"],
    treatment: "deduction_possible",
    atoLabel: "D5",
    category: "Potential work-related expense (category fallback)",
    confidence: "low",
    reason: "PocketSmith category fallback used; merchant text did not provide enough context."
  },
  {
    match: ["eating out", "restaurants", "alcohol", "groceries", "household", "entertainment"],
    treatment: "non_deductible",
    category: "Likely private expense (category fallback)",
    confidence: "low",
    reason: "PocketSmith category fallback used; merchant text did not provide enough context."
  }
];

const MERCHANT_STOPWORDS = new Set([
  "online",
  "payment",
  "payments",
  "receipt",
  "received",
  "thankyou",
  "thank",
  "date",
  "card",
  "debit",
  "credit",
  "from",
  "to",
  "ref",
  "help",
  "time",
  "eftpos",
  "purchase",
  "tap",
  "transfer",
  "funds",
  "internal",
  "a2a",
  "osko",
  "bpay",
  "bulk",
  "return",
  "returned",
  "bank",
  "transaction"
]);

const MERCHANT_LEGAL_SUFFIXES = new Set([
  "pty",
  "ltd",
  "limited",
  "co",
  "company",
  "inc",
  "llc",
  "plc",
  "corp",
  "corporation",
  "australia",
  "australian",
  "au",
  "group",
  "holdings",
  "trust",
  "trustee",
  "unit",
  "the"
]);

const MERCHANT_ALIAS_RULES = [
  { key: "transport for nsw", pattern: /\btransport\s*for\s*nsw\b/ },
  { key: "officeworks", pattern: /\bofficeworks\b/ },
  { key: "medrecruit", pattern: /\bmedrecruit\b/ },
  { key: "nsw health", pattern: /\bnsw\s*health\b|\bnswhealth\b/ },
  { key: "eyex australia", pattern: /\beyex\b/ },
  { key: "google", pattern: /\bgoogle\b|\bg\.co\b/ },
  { key: "paypal", pattern: /\bpaypal\b/ },
  { key: "tpg internet", pattern: /\btpg\b/ },
  { key: "ato", pattern: /\bato\b|\btax office\b/ },
  { key: "american express", pattern: /\bamerican express\b|\bamex\b/ },
  { key: "uber eats", pattern: /\buber\s*eats\b/ },
  { key: "uber", pattern: /\buber\b/ },
  { key: "xero", pattern: /\bxero\b/ },
  { key: "crypto.com", pattern: /\bcrypto\.?\s*com\b/ },
  { key: "spaceship", pattern: /\bspaceship\b/ },
  { key: "raiz", pattern: /\braiz\b/ },
  { key: "koinly", pattern: /\bkoinly\b/ },
  { key: "mda", pattern: /\bmda\b/ },
  { key: "ranzco", pattern: /\branzco\b/ }
];

const MERCHANT_INCOME_CATEGORIES = new Set([
  "employment_income",
  "health_employer",
  "government_employer",
  "staffing_agency"
]);

const MERCHANT_DEDUCTION_CATEGORIES = new Set([
  "office_supplies",
  "software_technology",
  "telecom_internet",
  "education_training",
  "professional_services",
  "tax_accounting"
]);

const MERCHANT_PRIVATE_CATEGORIES = new Set([
  "food_beverage",
  "grocery_retail",
  "general_retail",
  "entertainment_media",
  "personal_travel",
  "housing_private"
]);

const elements = {
  csvFile: document.getElementById("csvFile"),
  fileStatus: document.getElementById("fileStatus"),
  fySelect: document.getElementById("fySelect"),
  occupation: document.getElementById("occupation"),
  employers: document.getElementById("employers"),
  includePossible: document.getElementById("includePossible"),
  strictTransfers: document.getElementById("strictTransfers"),
  analyzeBtn: document.getElementById("analyzeBtn"),
  analysisStatus: document.getElementById("analysisStatus"),
  summaryPanel: document.getElementById("summaryPanel"),
  incomePanel: document.getElementById("incomePanel"),
  deductionPanel: document.getElementById("deductionPanel"),
  reviewPanel: document.getElementById("reviewPanel"),
  merchantPanel: document.getElementById("merchantPanel"),
  submissionPanel: document.getElementById("submissionPanel"),
  summaryCards: document.getElementById("summaryCards"),
  summaryMeta: document.getElementById("summaryMeta"),
  incomeTableBody: document.getElementById("incomeTableBody"),
  deductionTableBody: document.getElementById("deductionTableBody"),
  reviewTableBody: document.getElementById("reviewTableBody"),
  merchantTableBody: document.getElementById("merchantTableBody"),
  merchantStatus: document.getElementById("merchantStatus"),
  enrichMerchantsBtn: document.getElementById("enrichMerchantsBtn"),
  downloadMerchantsBtn: document.getElementById("downloadMerchantsBtn"),
  submissionText: document.getElementById("submissionText"),
  copySubmissionBtn: document.getElementById("copySubmissionBtn"),
  downloadSubmissionBtn: document.getElementById("downloadSubmissionBtn"),
  downloadClassifiedBtn: document.getElementById("downloadClassifiedBtn")
};

const state = {
  fileName: "",
  transactions: [],
  analysis: null,
  selectedFy: null,
  selectedOptions: null,
  currentFyTransactions: [],
  merchantGroups: [],
  merchantIntelByLookup: new Map()
};

bindEvents();

function bindEvents() {
  elements.csvFile.addEventListener("change", onFileSelected);
  elements.analyzeBtn.addEventListener("click", analyzeCurrentSelection);
  elements.enrichMerchantsBtn.addEventListener("click", enrichMerchantsForCurrentFy);
  elements.downloadMerchantsBtn.addEventListener("click", downloadMerchantsCsv);
  elements.copySubmissionBtn.addEventListener("click", copySubmission);
  elements.downloadSubmissionBtn.addEventListener("click", downloadSubmission);
  elements.downloadClassifiedBtn.addEventListener("click", downloadClassifiedCsv);
}

async function onFileSelected(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) {
    return;
  }

  setStatus(`Loading ${file.name} ...`);
  try {
    const text = await file.text();
    const records = parseCsv(text);
    const transactions = normalizeTransactions(records);

    if (!transactions.length) {
      throw new Error("No valid transaction rows were found in the CSV.");
    }

    state.fileName = file.name;
    state.transactions = transactions;
    state.analysis = null;
    state.selectedFy = null;
    state.selectedOptions = null;
    state.currentFyTransactions = [];
    state.merchantGroups = [];
    state.merchantIntelByLookup = new Map();

    elements.fileStatus.textContent = `${file.name} loaded (${transactions.length.toLocaleString()} transactions).`;
    populateFyOptions(transactions);
    clearResults();
    setMerchantStatus("Merchant enrichment has not been run.");
    elements.enrichMerchantsBtn.disabled = true;
    elements.downloadMerchantsBtn.disabled = true;
    setStatus("CSV loaded. Choose FY and click Analyze.");
  } catch (error) {
    elements.fileStatus.textContent = "Failed to load CSV file.";
    setStatus(error.message || "Failed to parse CSV.", true);
    elements.analyzeBtn.disabled = true;
    elements.fySelect.disabled = true;
  }
}

function analyzeCurrentSelection() {
  if (!state.transactions.length) {
    setStatus("Load a CSV file before analysis.", true);
    return;
  }

  const fy = Number(elements.fySelect.value);
  if (!Number.isFinite(fy)) {
    setStatus("Select a valid financial year.", true);
    return;
  }

  const transactions = state.transactions.filter((tx) => tx.fyEndYear === fy);
  if (!transactions.length) {
    setStatus(`No transactions found for FY${fy}.`, true);
    return;
  }

  const options = {
    includePossible: elements.includePossible.checked,
    strictTransfers: elements.strictTransfers.checked,
    occupation: elements.occupation.value.trim(),
    employerNames: parseEmployerNames(elements.employers.value)
  };

  state.selectedFy = fy;
  state.selectedOptions = options;
  state.currentFyTransactions = transactions;
  state.merchantGroups = buildMerchantGroups(transactions);

  setStatus(`Analyzing ${transactions.length.toLocaleString()} transactions for FY${fy} ...`);
  const analysis = analyzeTransactions(transactions, fy, options, state.merchantIntelByLookup, state.merchantGroups.length);
  state.analysis = analysis;
  renderAnalysis(analysis);

  elements.enrichMerchantsBtn.disabled = state.merchantGroups.length === 0;
  elements.downloadMerchantsBtn.disabled = state.merchantGroups.length === 0;
  renderMerchantRows(state.merchantGroups, state.merchantIntelByLookup);
  setMerchantStatus(
    state.merchantIntelByLookup.size
      ? `Merchant intelligence loaded for ${state.merchantIntelByLookup.size.toLocaleString()} merchants.`
      : `Ready to enrich ${state.merchantGroups.length.toLocaleString()} merchants for FY${fy}.`
  );
  setStatus(`Analysis complete for FY${fy}. Submission pack ready.`);
}

function analyzeTransactions(transactions, fy, options, merchantIntelByLookup, merchantCount) {
  const creditStats = buildCreditStats(transactions);
  const transferPairs = options.strictTransfers ? detectTransferPairs(transactions) : new Map();
  const reversalPairs = detectReversalPairs(transactions, transferPairs);
  const classified = transactions.map((tx) =>
    classifyTransaction(tx, {
      options,
      transferPairs,
      reversalPairs,
      creditStats,
      merchantIntelByLookup
    })
  );

  const summary = summariseClassifications(classified, options, merchantIntelByLookup, merchantCount);
  const markdown = buildSubmissionMarkdown({
    fileName: state.fileName,
    fy,
    options,
    summary
  });

  return {
    fileName: state.fileName,
    fy,
    options,
    transactions,
    classified,
    summary,
    markdown
  };
}

function classifyTransaction(tx, ctx) {
  const { transferPairs, reversalPairs, creditStats, options, merchantIntelByLookup } = ctx;
  const inTransferPair = transferPairs.has(tx.uid);
  const inReversalPair = reversalPairs.has(tx.uid);
  const merchantIntel = merchantIntelByLookup ? merchantIntelByLookup.get(tx.merchantLookupKey) : null;

  if (tx.explicitInternalLabel) {
    return result(tx, {
      treatment: "internal_transfer",
      taxCategory: "Internal transfer",
      confidence: "high",
      reason: "PocketSmith label marks this as internal."
    });
  }

  if (tx.explicitNonDeductibleLabel) {
    return result(tx, {
      treatment: "non_deductible",
      taxCategory: "Explicitly marked non-deductible",
      confidence: "high",
      reason: "PocketSmith label marks this transaction as non-deductible."
    });
  }

  if (tx.amount > 0) {
    const incomeMatch = classifyIncome(tx, creditStats, options.employerNames);
    if (incomeMatch) {
      return result(tx, incomeMatch);
    }

    if (inReversalPair || tx.refundKeyword) {
      return result(tx, {
        treatment: "excluded_refund",
        taxCategory: "Refund/reversal (excluded)",
        confidence: inReversalPair ? "high" : "medium",
        reason: inReversalPair ? "Matched equal opposite transaction, likely reversal/refund." : "Merchant text suggests refund/reversal."
      });
    }

    if (inTransferPair || tx.transferKeyword) {
      return result(tx, {
        treatment: "internal_transfer",
        taxCategory: "Likely internal transfer",
        confidence: inTransferPair ? "high" : "medium",
        reason: "Merchant text indicates transfer movement between accounts."
      });
    }

    const merchantIntelMatch = classifyFromMerchantIntel(tx, merchantIntel);
    if (merchantIntelMatch) {
      return result(tx, merchantIntelMatch);
    }

    if (containsAny(tx.textMerchant, ["ato"])) {
      return result(tx, {
        treatment: "income_review",
        taxCategory: "ATO credit (review)",
        confidence: "medium",
        reason: "ATO-related credit detected. Confirm if this is a tax refund/non-assessable amount."
      });
    }

    return result(tx, {
      treatment: "income_review",
      taxCategory: "Other credit (review)",
      confidence: "low",
      reason: "Credit not confidently classified from merchant text."
    });
  }

  if (inTransferPair || tx.transferKeyword) {
    return result(tx, {
      treatment: "internal_transfer",
      taxCategory: "Likely internal transfer",
      confidence: inTransferPair ? "high" : "medium",
      reason: "Merchant text indicates transfer movement between accounts."
    });
  }

  if (inReversalPair || tx.refundKeyword) {
    return result(tx, {
      treatment: "excluded_refund",
      taxCategory: "Reversal/refund (excluded)",
      confidence: inReversalPair ? "high" : "medium",
      reason: inReversalPair ? "Matched equal opposite transaction, likely reversal/refund." : "Merchant text suggests a refund/reversal flow."
    });
  }

  if (tx.explicitDeductibleLabel) {
    const merchantRule = matchDeductionRule(tx);
    if (merchantRule) {
      return result(tx, {
        treatment: "deduction_likely",
        atoLabel: merchantRule.atoLabel,
        taxCategory: merchantRule.category,
        confidence: "high",
        reason: `PocketSmith tax label + merchant pattern. ${merchantRule.note}`
      });
    }

    return result(tx, {
      treatment: "deduction_likely",
      atoLabel: "D5",
      taxCategory: "Other work-related expenses (tagged)",
      confidence: "high",
      reason: "PocketSmith tax label applied; merchant text did not map to a more specific deduction category."
    });
  }

  if (tx.taxPaymentKeyword && !tx.taxAgentKeyword) {
    return result(tx, {
      treatment: "non_deductible",
      taxCategory: "Income tax or ATO payment",
      confidence: "high",
      reason: "Merchant text indicates tax liability payment, typically not deductible."
    });
  }

  const merchantIntelMatch = classifyFromMerchantIntel(tx, merchantIntel);
  if (merchantIntelMatch) {
    return result(tx, merchantIntelMatch);
  }

  const deductionRule = matchDeductionRule(tx);
  if (deductionRule) {
    return result(tx, {
      treatment: deductionRule.defaultTreatment,
      atoLabel: deductionRule.atoLabel,
      taxCategory: deductionRule.category,
      confidence: deductionRule.confidence,
      reason: `Matched from merchant text. ${deductionRule.note}`
    });
  }

  const nonDeductibleRule = matchNonDeductibleRule(tx);
  if (nonDeductibleRule) {
    return result(tx, {
      treatment: "non_deductible",
      taxCategory: nonDeductibleRule.category,
      confidence: "medium",
      reason: "Merchant text strongly suggests private/personal spending."
    });
  }

  const fallback = classifyByCategoryFallback(tx);
  if (fallback) {
    return result(tx, fallback);
  }

  return result(tx, {
    treatment: "review",
    taxCategory: "Expense needs review",
    confidence: "low",
    reason: "Not confidently classifiable from merchant text; accountant review required."
  });
}

function classifyIncome(tx, creditStats, employerNames) {
  if (tx.interestKeyword) {
    return {
      treatment: "income_assessable",
      taxCategory: "Bank interest",
      confidence: "high",
      reason: "Merchant text indicates interest income."
    };
  }

  if (containsAny(tx.textMerchant, INVESTMENT_INCOME_KEYWORDS)) {
    return {
      treatment: "income_assessable",
      taxCategory: "Investment income",
      confidence: "medium",
      reason: "Merchant text indicates distribution/dividend income."
    };
  }

  let salaryScore = 0;
  if (tx.salaryKeyword) {
    salaryScore += 3;
  }
  if (hasAnyLabel(tx.labels, ["locum", "nswhealth", "salary", "wages"])) {
    salaryScore += 3;
  }
  if (matchesEmployerName(tx, employerNames)) {
    salaryScore += 3;
  }

  const stat = creditStats.get(tx.recurrenceKey);
  if (stat && stat.count >= 3) {
    salaryScore += 1;
  }
  if (tx.absAmount >= 500) {
    salaryScore += 1;
  }
  if (tx.transferKeyword) {
    salaryScore -= 2;
  }

  if (salaryScore >= 4) {
    return {
      treatment: "income_assessable",
      taxCategory: "Salary and wages",
      confidence: salaryScore >= 6 ? "high" : "medium",
      reason: "Merchant text and recurrence suggest employment income."
    };
  }

  return null;
}

function classifyFromMerchantIntel(tx, intel) {
  if (!intel || !intel.businessCategory) {
    return null;
  }

  const category = String(intel.businessCategory || "").trim().toLowerCase();
  if (!category || category === "unknown") {
    return null;
  }

  const confidence = normalizeConfidence(intel.classificationConfidence);
  const intelContext = [
    intel.businessType ? `business type: ${intel.businessType}` : "",
    intel.abn ? `ABN: ${intel.abn}` : "",
    intel.mainPlaceOfBusiness ? `main place: ${intel.mainPlaceOfBusiness}` : ""
  ]
    .filter(Boolean)
    .join("; ");
  const reasonPrefix = intelContext ? `Merchant intelligence (${intelContext}). ` : "Merchant intelligence applied. ";

  if (tx.amount > 0) {
    if (category === "banking_event") {
      return {
        treatment: "internal_transfer",
        taxCategory: "Internal transfer / banking event",
        confidence: bumpConfidence(confidence, "medium"),
        reason: `${reasonPrefix}Credit treated as non-assessable banking movement.`
      };
    }

    if (MERCHANT_INCOME_CATEGORIES.has(category)) {
      return {
        treatment: "income_assessable",
        taxCategory: "Salary and wages",
        confidence: bumpConfidence(confidence, "medium"),
        reason: `${reasonPrefix}Credit treated as likely employment income.`
      };
    }

    if (category === "banking_financial" && tx.interestKeyword) {
      return {
        treatment: "income_assessable",
        taxCategory: "Bank interest",
        confidence: bumpConfidence(confidence, "medium"),
        reason: `${reasonPrefix}Credit treated as likely interest/financial income.`
      };
    }

    return null;
  }

  if (category === "banking_event") {
    return {
      treatment: "internal_transfer",
      taxCategory: "Internal transfer / banking event",
      confidence: bumpConfidence(confidence, "medium"),
      reason: `${reasonPrefix}Debit treated as non-deductible banking movement.`
    };
  }

  if (category === "tax_accounting") {
    return {
      treatment: "deduction_likely",
      atoLabel: "D10",
      taxCategory: "Cost of managing tax affairs",
      confidence: bumpConfidence(confidence, "medium"),
      reason: `${reasonPrefix}Merchant category suggests tax/accounting service expense.`
    };
  }

  if (MERCHANT_DEDUCTION_CATEGORIES.has(category)) {
    return {
      treatment: "deduction_possible",
      atoLabel: "D5",
      taxCategory: "Potential work-related expense (merchant intelligence)",
      confidence,
      reason: `${reasonPrefix}Merchant category suggests potential work-related expense requiring apportionment review.`
    };
  }

  if (MERCHANT_PRIVATE_CATEGORIES.has(category)) {
    return {
      treatment: "non_deductible",
      taxCategory: "Likely private expense (merchant intelligence)",
      confidence: bumpConfidence(confidence, "medium"),
      reason: `${reasonPrefix}Merchant category suggests private/personal spending.`
    };
  }

  return null;
}

function normalizeConfidence(level) {
  if (level === "high" || level === "medium" || level === "low") {
    return level;
  }
  return "low";
}

function bumpConfidence(level, minimum) {
  const order = ["low", "medium", "high"];
  return order.indexOf(level) >= order.indexOf(minimum) ? level : minimum;
}

function matchDeductionRule(tx) {
  for (const rule of DEDUCTION_RULES) {
    if (containsAny(`${tx.textMerchant} ${tx.textMeta}`, rule.keywords)) {
      return rule;
    }
  }
  return null;
}

function matchNonDeductibleRule(tx) {
  for (const rule of NON_DEDUCTIBLE_RULES) {
    if (containsAny(tx.textMerchant, rule.keywords)) {
      return rule;
    }
  }
  return null;
}

function classifyByCategoryFallback(tx) {
  const categoryText = normalizeText(`${tx.category} ${tx.parentCategory}`);
  if (!categoryText) {
    return null;
  }

  for (const fallback of CATEGORY_FALLBACK_RULES) {
    if (containsAny(categoryText, fallback.match)) {
      return {
        treatment: fallback.treatment,
        atoLabel: fallback.atoLabel || "",
        taxCategory: fallback.category,
        confidence: fallback.confidence,
        reason: fallback.reason
      };
    }
  }
  return null;
}

function summariseClassifications(classified, options, merchantIntelByLookup, merchantCount) {
  const incomeMap = new Map();
  const deductionMap = new Map();
  const summary = {
    transactionCount: classified.length,
    merchantCount: merchantCount || 0,
    enrichedMerchants: 0,
    resolvedMerchants: 0,
    assessableIncome: 0,
    likelyDeductions: 0,
    possibleDeductions: 0,
    includedDeductions: 0,
    nonDeductible: 0,
    internalTransfers: 0,
    excludedRefunds: 0,
    reviewCount: 0,
    reviewAmount: 0,
    incomeRows: [],
    deductionRows: [],
    reviewRows: []
  };

  for (const item of classified) {
    const amount = item.tx.absAmount;

    if (item.treatment === "income_assessable") {
      summary.assessableIncome += item.tx.amount;
      accumulate(incomeMap, item.taxCategory, item.tx.amount, item.confidence, "", item.taxCategory, item.treatment);
      continue;
    }

    if (item.treatment === "deduction_likely") {
      summary.likelyDeductions += amount;
      summary.includedDeductions += amount;
      accumulate(deductionMap, `${item.atoLabel}|${item.taxCategory}|${item.treatment}`, amount, item.confidence, item.atoLabel, item.taxCategory, item.treatment);
      continue;
    }

    if (item.treatment === "deduction_possible") {
      summary.possibleDeductions += amount;
      if (options.includePossible) {
        summary.includedDeductions += amount;
      }
      accumulate(deductionMap, `${item.atoLabel}|${item.taxCategory}|${item.treatment}`, amount, item.confidence, item.atoLabel, item.taxCategory, item.treatment);
      summary.reviewRows.push(item);
      continue;
    }

    if (item.treatment === "internal_transfer") {
      summary.internalTransfers += amount;
      continue;
    }

    if (item.treatment === "excluded_refund") {
      summary.excludedRefunds += amount;
      continue;
    }

    if (item.treatment === "non_deductible") {
      summary.nonDeductible += amount;
      continue;
    }

    if (item.treatment === "income_review" || item.treatment === "review") {
      summary.reviewRows.push(item);
      summary.reviewCount += 1;
      summary.reviewAmount += amount;
    }
  }

  summary.incomeRows = Array.from(incomeMap.values()).sort((a, b) => b.amount - a.amount);
  summary.deductionRows = Array.from(deductionMap.values()).sort((a, b) => {
    if (a.treatment !== b.treatment) {
      return a.treatment === "deduction_likely" ? -1 : 1;
    }
    return b.amount - a.amount;
  });
  summary.reviewRows = summary.reviewRows.sort((a, b) => b.tx.absAmount - a.tx.absAmount).slice(0, 80);

  const merchantKeys = new Set(classified.map((item) => item.tx.merchantLookupKey).filter(Boolean));
  summary.merchantCount = summary.merchantCount || merchantKeys.size;
  if (merchantIntelByLookup) {
    for (const key of merchantKeys) {
      const intel = merchantIntelByLookup.get(key);
      if (!intel) {
        continue;
      }
      summary.enrichedMerchants += 1;
      if (intel.businessCategory && intel.businessCategory !== "unknown") {
        summary.resolvedMerchants += 1;
      }
    }
  }
  return summary;
}

function accumulate(map, key, amount, confidence, atoLabel, category, treatment) {
  if (!map.has(key)) {
    map.set(key, {
      key,
      atoLabel,
      category,
      treatment,
      count: 0,
      amount: 0,
      high: 0,
      medium: 0,
      low: 0
    });
  }
  const row = map.get(key);
  row.count += 1;
  row.amount += amount;
  row[confidence] = (row[confidence] || 0) + 1;
}

function dominantConfidence(row) {
  if ((row.high || 0) >= (row.medium || 0) && (row.high || 0) >= (row.low || 0)) {
    return "high";
  }
  if ((row.medium || 0) >= (row.low || 0)) {
    return "medium";
  }
  return "low";
}

function buildSubmissionMarkdown({ fileName, fy, options, summary }) {
  const nowText = formatDateTimeSydney(new Date());
  const periodText = fyPeriodText(fy);
  const incomeLines = summary.incomeRows.length
    ? summary.incomeRows.map((row) => `| ${escapeMarkdown(row.category)} | ${row.count} | ${formatAud(row.amount)} | ${dominantConfidence(row)} |`).join("\n")
    : "| None identified | 0 | $0.00 | low |";

  const likelyRows = summary.deductionRows.filter((row) => row.treatment === "deduction_likely");
  const possibleRows = summary.deductionRows.filter((row) => row.treatment === "deduction_possible");

  const likelyLines = likelyRows.length
    ? likelyRows.map((row) => `| ${row.atoLabel || "-"} | ${escapeMarkdown(row.category)} | ${row.count} | ${formatAud(row.amount)} | ${dominantConfidence(row)} |`).join("\n")
    : "| - | None identified | 0 | $0.00 | low |";

  const possibleLines = possibleRows.length
    ? possibleRows.map((row) => `| ${row.atoLabel || "-"} | ${escapeMarkdown(row.category)} | ${row.count} | ${formatAud(row.amount)} | ${dominantConfidence(row)} |`).join("\n")
    : "| - | None identified | 0 | $0.00 | low |";

  const reviewLines = summary.reviewRows.length
    ? summary.reviewRows
        .slice(0, 20)
        .map(
          (item) =>
            `| ${item.tx.dateText} | ${escapeMarkdown(item.tx.merchant)} | ${formatAud(item.tx.absAmount)} | ${escapeMarkdown(displayTreatment(item.treatment))} | ${escapeMarkdown(item.reason)} |`
        )
        .join("\n")
    : "| - | - | $0.00 | - | No priority review items in top list |";

  return [
    "# Tax Return Submission Pack (Australia - Employee)",
    "",
    `Generated: ${nowText}`,
    `Source CSV: ${fileName || "(unknown)"}`,
    `Financial Year: FY${fy} (${periodText})`,
    options.occupation ? `Occupation: ${options.occupation}` : "Occupation: (not supplied)",
    options.employerNames.length ? `Known employers used in matching: ${options.employerNames.join(", ")}` : "Known employers used in matching: (none supplied)",
    "",
    "## Method Notes",
    "- Merchant transaction names are the primary classification signal.",
    "- Merchant web enrichment can be run in-app (ABR/ABN lookup + web search) to classify each merchant and guide tax treatment.",
    "- PocketSmith categories are used only as low-confidence fallback when merchant text is inconclusive.",
    "- Internal transfers and mirrored reversals are excluded from income/deduction totals.",
    "- This is a preparatory draft for accountant review, not a lodgement-ready declaration.",
    "",
    "## Income Schedule",
    "| Category | Transactions | Amount (AUD) | Confidence |",
    "| --- | ---: | ---: | --- |",
    incomeLines,
    "",
    `Total assessable income identified: **${formatAud(summary.assessableIncome)}**`,
    "",
    "## Deduction Schedule - Likely",
    "| ATO Label | Category | Transactions | Amount (AUD) | Confidence |",
    "| --- | --- | ---: | ---: | --- |",
    likelyLines,
    "",
    `Total likely deductions: **${formatAud(summary.likelyDeductions)}**`,
    "",
    "## Deduction Schedule - Possible (Review Required)",
    "| ATO Label | Category | Transactions | Amount (AUD) | Confidence |",
    "| --- | --- | ---: | ---: | --- |",
    possibleLines,
    "",
    `Total possible deductions: **${formatAud(summary.possibleDeductions)}**`,
    `Included in draft total: **${formatAud(summary.includedDeductions)}** (${options.includePossible ? "includes possible deductions" : "likely deductions only"})`,
    "",
    "## Exclusions and Non-Deductible",
    `- Internal transfers excluded: ${formatAud(summary.internalTransfers)}`,
    `- Reversals/refunds excluded: ${formatAud(summary.excludedRefunds)}`,
    `- Non-deductible/private expenses identified: ${formatAud(summary.nonDeductible)}`,
    "",
    "## Merchant Intelligence Coverage",
    `- Unique merchants in FY: ${summary.merchantCount.toLocaleString()}`,
    `- Merchants web-enriched: ${summary.enrichedMerchants.toLocaleString()}`,
    `- Merchants with resolved business category: ${summary.resolvedMerchants.toLocaleString()}`,
    "",
    "## Priority Accountant Review Items",
    "| Date | Merchant | Amount (AUD) | Suggested Treatment | Reason |",
    "| --- | --- | ---: | --- | --- |",
    reviewLines,
    "",
    "## Evidence Checklist",
    "- Receipts/invoices for all claimed work-related expenses.",
    "- Basis for any private/work apportionment (internet, phone, subscriptions).",
    "- Work travel substantiation (purpose, diary/logbook where required).",
    "- Tax agent invoices and donation receipts (DGR confirmation).",
    "- Clarification for all items listed in the review table above.",
    "",
    "## Accountant Questions",
    "1. Confirm whether any uploaded accounts are business/non-personal accounts and adjust inclusions.",
    "2. Confirm treatment of ATO-related credits/debits and any prior-year adjustments.",
    "3. Validate any potential investment-related amounts that may not belong in employee deductions.",
    "4. Confirm work-related necessity and apportionment percentages for possible deductions."
  ].join("\n");
}

function detectTransferPairs(transactions) {
  const byAmount = groupByAbsAmount(transactions);
  const pairs = new Map();

  for (const list of byAmount.values()) {
    list.sort((a, b) => a.dateValue - b.dateValue);
    for (let i = 0; i < list.length; i += 1) {
      const tx = list[i];
      if (pairs.has(tx.uid)) {
        continue;
      }

      let bestMatch = null;
      let bestScore = 0;
      for (let j = i + 1; j < list.length; j += 1) {
        const candidate = list[j];
        if (pairs.has(candidate.uid)) {
          continue;
        }
        if (Math.sign(tx.amount) === Math.sign(candidate.amount)) {
          continue;
        }

        const dayGap = Math.abs(candidate.dateValue - tx.dateValue) / DAY_MS;
        if (dayGap > 3) {
          break;
        }

        const score = transferPairScore(tx, candidate);
        if (score > bestScore) {
          bestScore = score;
          bestMatch = candidate;
        }
      }

      if (bestMatch && bestScore >= 4) {
        pairs.set(tx.uid, bestMatch.uid);
        pairs.set(bestMatch.uid, tx.uid);
      }
    }
  }
  return pairs;
}

function transferPairScore(a, b) {
  let score = 0;
  if (a.transferKeyword || b.transferKeyword) {
    score += 3;
  }
  if (a.explicitInternalLabel || b.explicitInternalLabel) {
    score += 3;
  }
  if (a.account && b.account && a.account !== b.account) {
    score += 2;
  }
  if (Math.abs(a.dateValue - b.dateValue) <= DAY_MS) {
    score += 1;
  }
  if (a.recurrenceKey && b.recurrenceKey && a.recurrenceKey === b.recurrenceKey && (a.transferKeyword || b.transferKeyword)) {
    score += 1;
  }
  return score;
}

function detectReversalPairs(transactions, transferPairs) {
  const byAmount = groupByAbsAmount(transactions.filter((tx) => !transferPairs.has(tx.uid)));
  const pairs = new Map();

  for (const list of byAmount.values()) {
    list.sort((a, b) => a.dateValue - b.dateValue);
    for (let i = 0; i < list.length; i += 1) {
      const tx = list[i];
      if (pairs.has(tx.uid) || tx.salaryKeyword) {
        continue;
      }

      let bestMatch = null;
      let bestScore = 0;
      for (let j = i + 1; j < list.length; j += 1) {
        const candidate = list[j];
        if (pairs.has(candidate.uid) || candidate.salaryKeyword) {
          continue;
        }
        if (Math.sign(tx.amount) === Math.sign(candidate.amount)) {
          continue;
        }

        const dayGap = Math.abs(candidate.dateValue - tx.dateValue) / DAY_MS;
        if (dayGap > 45) {
          break;
        }

        const similarity = merchantSimilarity(tx.recurrenceKey, candidate.recurrenceKey);
        const score = similarity + (tx.refundKeyword || candidate.refundKeyword ? 0.2 : 0);
        if (score > bestScore) {
          bestScore = score;
          bestMatch = candidate;
        }
      }

      if (bestMatch && bestScore >= 0.62) {
        pairs.set(tx.uid, bestMatch.uid);
        pairs.set(bestMatch.uid, tx.uid);
      }
    }
  }
  return pairs;
}

function groupByAbsAmount(transactions) {
  const map = new Map();
  for (const tx of transactions) {
    const key = tx.absAmount.toFixed(2);
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(tx);
  }
  return map;
}

function buildCreditStats(transactions) {
  const map = new Map();
  for (const tx of transactions) {
    if (tx.amount <= 0) {
      continue;
    }
    const key = tx.recurrenceKey;
    if (!key) {
      continue;
    }
    if (!map.has(key)) {
      map.set(key, { count: 0, total: 0 });
    }
    const stat = map.get(key);
    stat.count += 1;
    stat.total += tx.amount;
  }
  return map;
}

function parseEmployerNames(input) {
  return String(input || "")
    .split(",")
    .map((item) => normalizeText(item))
    .filter(Boolean);
}

function matchesEmployerName(tx, employerNames) {
  if (!employerNames.length) {
    return false;
  }
  return employerNames.some((name) => tx.textMerchant.includes(name));
}

function buildMerchantGroups(transactions) {
  const map = new Map();

  for (const tx of transactions) {
    const lookupKey = tx.merchantLookupKey || deriveMerchantLookupKey(tx.merchant);
    if (!lookupKey) {
      continue;
    }

    if (!map.has(lookupKey)) {
      map.set(lookupKey, {
        lookupKey,
        sampleMerchant: tx.merchant,
        transactionCount: 0,
        totalAbsAmount: 0,
        merchantNameCounts: new Map()
      });
    }

    const group = map.get(lookupKey);
    group.transactionCount += 1;
    group.totalAbsAmount += tx.absAmount;
    group.merchantNameCounts.set(tx.merchant, (group.merchantNameCounts.get(tx.merchant) || 0) + 1);

    if (!group.sampleMerchant || group.sampleMerchant.toLowerCase().includes("unknown")) {
      group.sampleMerchant = tx.merchant;
    }
  }

  const groups = Array.from(map.values());
  for (const group of groups) {
    let bestName = group.sampleMerchant;
    let bestCount = 0;
    for (const [name, count] of group.merchantNameCounts.entries()) {
      if (count > bestCount) {
        bestCount = count;
        bestName = name;
      }
    }
    group.sampleMerchant = bestName;
    delete group.merchantNameCounts;
  }

  return groups.sort((a, b) => b.totalAbsAmount - a.totalAbsAmount);
}

function deriveMerchantLookupKey(rawMerchant) {
  const normalized = normalizeText(rawMerchant || "");
  if (!normalized) {
    return "";
  }

  let cleaned = normalized;
  cleaned = cleaned.replace(/^transportfornsw/, "transport for nsw");
  cleaned = cleaned.replace(/^paypal\s+\*?\s*/, "");
  cleaned = cleaned.replace(/^google\s+\*?\s*/, "google ");
  cleaned = cleaned.replace(/\\n/g, " ");
  cleaned = cleaned.replace(/\s+/g, " ").trim();

  for (const alias of MERCHANT_ALIAS_RULES) {
    if (alias.pattern.test(cleaned)) {
      return alias.key;
    }
  }

  if (isLikelyInternalTransferMerchant(cleaned)) {
    return "internal transfer";
  }

  let tokens = cleaned
    .split(" ")
    .map((token) => normalizeMerchantToken(token))
    .filter(Boolean)
    .filter((token) => !isMerchantNoiseToken(token));

  if (!tokens.length) {
    return normalized.slice(0, 80);
  }

  tokens = dedupeTokens(tokens);

  if (tokens.includes("eyex")) {
    return "eyex australia";
  }
  if (tokens.includes("officeworks")) {
    return "officeworks";
  }
  if (tokens.includes("medrecruit")) {
    return "medrecruit";
  }
  if (tokens.includes("google")) {
    return "google";
  }
  if (tokens.includes("transport") && tokens.includes("nsw")) {
    return "transport for nsw";
  }
  if (tokens.includes("nsw") && tokens.includes("health")) {
    return "nsw health";
  }
  if (tokens.includes("tpg")) {
    return "tpg internet";
  }

  const strongTokens = tokens.filter((token) => !MERCHANT_LEGAL_SUFFIXES.has(token));
  const chosen = strongTokens.length ? strongTokens : tokens;
  if (!chosen.length) {
    return normalized.slice(0, 80);
  }

  if (chosen.length === 1) {
    return chosen[0];
  }

  return chosen.slice(0, 2).join(" ");
}

function normalizeMerchantToken(token) {
  let value = String(token || "").trim().toLowerCase();
  if (!value) {
    return "";
  }

  if (value === "austral") {
    return "australia";
  }
  if (value === "nswhealth") {
    return "nsw health";
  }
  if (value === "transportfornsw") {
    return "transport for nsw";
  }

  return value;
}

function isMerchantNoiseToken(token) {
  const value = String(token || "").trim();
  if (!value) {
    return true;
  }
  if (MERCHANT_STOPWORDS.has(value)) {
    return true;
  }
  if (/^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)$/.test(value)) {
    return true;
  }
  if (/^[a-z]*\d+[a-z0-9]*$/.test(value)) {
    return true;
  }
  if (/^x{2,}\d*$/i.test(value)) {
    return true;
  }
  if (/^\d+$/.test(value)) {
    return true;
  }
  if (value.length <= 1) {
    return true;
  }
  return false;
}

function dedupeTokens(tokens) {
  const out = [];
  const seen = new Set();
  for (const token of tokens) {
    if (!token || seen.has(token)) {
      continue;
    }
    seen.add(token);
    out.push(token);
  }
  return out;
}

function isLikelyInternalTransferMerchant(text) {
  const value = normalizeText(text || "");
  if (!value) {
    return false;
  }

  const hasTransferSignal = /\b(a2a|transfer|osko|round up|internal|savings maximiser|orange everyday|offset)\b/.test(value);
  if (!hasTransferSignal) {
    return false;
  }

  const hasKnownExternalMerchant = /\b(officeworks|medrecruit|google|tpg|transport for nsw|uber|paypal|xero|ato|eyex|nsw health|spaceship|raiz|koinly|mda|ranzco|woolworths|coles|aldi)\b/.test(
    value
  );
  return !hasKnownExternalMerchant;
}

function isSyntheticMerchantLookupKey(key) {
  const value = normalizeText(key || "");
  if (!value) {
    return false;
  }

  if (value === "internal transfer") {
    return true;
  }
  if (/\bloan drawdown\b/.test(value)) {
    return true;
  }
  if (/\bloan repayment\b/.test(value)) {
    return true;
  }
  if (value === "withdrawal") {
    return true;
  }
  return false;
}

function populateFyOptions(transactions) {
  const fyValues = Array.from(new Set(transactions.map((tx) => tx.fyEndYear))).sort((a, b) => a - b);
  elements.fySelect.innerHTML = "";

  for (const fy of fyValues) {
    const option = document.createElement("option");
    option.value = String(fy);
    option.textContent = `FY${fy} (${fyPeriodText(fy)})`;
    elements.fySelect.appendChild(option);
  }

  const defaultFy = chooseDefaultFy(fyValues);
  elements.fySelect.value = String(defaultFy);
  elements.fySelect.disabled = false;
  elements.analyzeBtn.disabled = false;
}

function chooseDefaultFy(fyValues) {
  if (!fyValues.length) {
    return "";
  }

  const latestEnded = latestEndedFySydney();
  const notFuture = fyValues.filter((fy) => fy <= latestEnded);
  if (notFuture.length) {
    return notFuture[notFuture.length - 1];
  }
  return fyValues[fyValues.length - 1];
}

function latestEndedFySydney() {
  const now = new Date();
  const parts = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Sydney",
    year: "numeric",
    month: "numeric"
  }).formatToParts(now);

  const year = Number(parts.find((part) => part.type === "year").value);
  const month = Number(parts.find((part) => part.type === "month").value);

  return month >= 7 ? year : year - 1;
}

function fyPeriodText(fy) {
  return `1 Jul ${fy - 1} - 30 Jun ${fy}`;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === "\"") {
        if (text[i + 1] === "\"") {
          cell += "\"";
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        cell += ch;
      }
      continue;
    }

    if (ch === "\"") {
      inQuotes = true;
      continue;
    }
    if (ch === ",") {
      row.push(cell);
      cell = "";
      continue;
    }
    if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }
    if (ch === "\r") {
      continue;
    }
    cell += ch;
  }

  if (cell.length > 0 || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }

  const filtered = rows.filter((r) => r.some((field) => String(field).trim() !== ""));
  if (filtered.length < 2) {
    return [];
  }

  const headers = filtered[0].map((h) => String(h).trim());
  const records = [];

  for (let i = 1; i < filtered.length; i += 1) {
    const values = filtered[i];
    const obj = {};
    headers.forEach((header, idx) => {
      obj[header] = values[idx] !== undefined ? values[idx] : "";
    });
    records.push(obj);
  }

  return records;
}

function normalizeTransactions(records) {
  const transactions = [];

  for (let i = 0; i < records.length; i += 1) {
    const record = records[i];
    const dateText = String(record.Date || "").trim();
    const dateParts = parseDateParts(dateText);
    if (!dateParts) {
      continue;
    }

    const amount = parseAmount(record.Amount);
    if (!Number.isFinite(amount)) {
      continue;
    }

    const merchant = String(record.Merchant || "").trim() || "Unknown merchant";
    const category = String(record.Category || "").trim();
    const parentCategory = String(record["Parent Categories"] || "").trim();
    const labels = extractLabels(record.Labels, record.Note);

    const tx = {
      uid: String(record.ID || "").trim() || `row-${i + 2}`,
      sourceRow: i + 2,
      dateText,
      dateValue: dateParts.dateValue,
      fyEndYear: dateParts.month >= 7 ? dateParts.year + 1 : dateParts.year,
      merchant,
      merchantChangedFrom: String(record["Merchant Changed From"] || "").trim(),
      amount,
      absAmount: Math.abs(amount),
      currency: String(record.Currency || CURRENCY).trim().toUpperCase(),
      transactionType: String(record["Transaction Type"] || "").trim().toLowerCase(),
      account: String(record.Account || "").trim(),
      category,
      parentCategory,
      labels,
      memo: String(record.Memo || "").trim(),
      note: String(record.Note || "").trim(),
      pocketId: String(record.ID || "").trim(),
      bank: String(record.Bank || "").trim(),
      accountNumber: String(record["Account Number"] || "").trim()
    };

    tx.textMerchant = normalizeText(`${tx.merchant} ${tx.merchantChangedFrom}`);
    tx.textMeta = normalizeText(`${tx.memo} ${tx.note} ${tx.labels.join(" ")}`);
    tx.textCategory = normalizeText(`${tx.category} ${tx.parentCategory}`);
    tx.merchantLookupKey = deriveMerchantLookupKey(tx.merchant);

    tx.transferKeyword = containsAny(`${tx.textMerchant} ${tx.textMeta}`, TRANSFER_KEYWORDS_PRIMARY) || hasAnyLabel(tx.labels, ["internal"]);
    tx.salaryKeyword = containsAny(`${tx.textMerchant} ${tx.textMeta}`, SALARY_KEYWORDS);
    tx.interestKeyword = containsAny(`${tx.textMerchant} ${tx.textMeta}`, INTEREST_KEYWORDS);
    tx.refundKeyword = containsAny(`${tx.textMerchant} ${tx.textMeta}`, REFUND_KEYWORDS);
    tx.taxPaymentKeyword = containsAny(`${tx.textMerchant} ${tx.textMeta}`, TAX_PAYMENT_KEYWORDS);
    tx.taxAgentKeyword = containsAny(`${tx.textMerchant} ${tx.textMeta}`, TAX_AGENT_KEYWORDS);

    tx.explicitInternalLabel = hasAnyLabel(tx.labels, ["internal"]);
    tx.explicitDeductibleLabel = hasAnyLabel(tx.labels, ["tax", "deductable", "deductible"]);
    tx.explicitNonDeductibleLabel = hasAnyLabel(tx.labels, ["nondeductable", "non-deductable", "non-deductible", "private"]);
    tx.recurrenceKey = makeMerchantKey(tx.textMerchant);

    transactions.push(tx);
  }

  transactions.sort((a, b) => a.dateValue - b.dateValue);
  return transactions;
}

function parseDateParts(value) {
  const match = /^([0-9]{4})-([0-9]{2})-([0-9]{2})$/.exec(value);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!year || !month || !day) {
    return null;
  }

  return {
    year,
    month,
    day,
    dateValue: Date.UTC(year, month - 1, day)
  };
}

function parseAmount(value) {
  const cleaned = String(value || "")
    .replace(/,/g, "")
    .replace(/\$/g, "")
    .trim();
  return Number(cleaned);
}

function extractLabels(labelsValue, noteValue) {
  const tokens = [];

  const labelParts = String(labelsValue || "")
    .split(/[\s,;|]+/)
    .map((item) => normalizeLabel(item))
    .filter(Boolean);
  tokens.push(...labelParts);

  const hashMatches = String(noteValue || "").match(/#[A-Za-z0-9._-]+/g) || [];
  for (const token of hashMatches) {
    const clean = normalizeLabel(token.replace(/^#/, ""));
    if (clean) {
      tokens.push(clean);
    }
  }

  return Array.from(new Set(tokens));
}

function normalizeLabel(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/^#+/, "")
    .replace(/[^a-z0-9._-]/g, "");
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[\\/\-_*]+/g, " ")
    .replace(/[^a-z0-9.& ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function containsAny(text, keywords) {
  const haystack = String(text || "");
  if (!haystack) {
    return false;
  }
  return keywords.some((keyword) => haystack.includes(keyword));
}

function hasAnyLabel(labels, wanted) {
  if (!labels.length) {
    return false;
  }
  return wanted.some((name) => labels.includes(name));
}

function makeMerchantKey(textMerchant) {
  const withoutNoise = String(textMerchant || "")
    .replace(/\b(online|payment|receipt|thankyou|thank|from|to|card|ref|xxxx|x{2,}|transfer|a2a|osko)\b/g, " ")
    .replace(/\d+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (withoutNoise) {
    return withoutNoise;
  }

  return String(textMerchant || "").trim();
}

function merchantSimilarity(a, b) {
  const first = String(a || "").trim();
  const second = String(b || "").trim();
  if (!first || !second) {
    return 0;
  }
  if (first === second) {
    return 1;
  }

  const aTokens = new Set(first.split(" ").filter((token) => token.length > 2));
  const bTokens = new Set(second.split(" ").filter((token) => token.length > 2));

  if (!aTokens.size || !bTokens.size) {
    return 0;
  }

  let intersection = 0;
  for (const token of aTokens) {
    if (bTokens.has(token)) {
      intersection += 1;
    }
  }
  const union = new Set([...aTokens, ...bTokens]).size;
  return union ? intersection / union : 0;
}

function result(tx, partial) {
  return {
    tx,
    treatment: partial.treatment,
    taxCategory: partial.taxCategory || "",
    atoLabel: partial.atoLabel || "",
    confidence: partial.confidence || "low",
    reason: partial.reason || ""
  };
}

function renderAnalysis(analysis) {
  const { summary, fy, options } = analysis;

  elements.summaryPanel.classList.remove("hidden");
  elements.incomePanel.classList.remove("hidden");
  elements.deductionPanel.classList.remove("hidden");
  elements.reviewPanel.classList.remove("hidden");
  elements.merchantPanel.classList.remove("hidden");
  elements.submissionPanel.classList.remove("hidden");

  const cards = [
    card("Transactions", summary.transactionCount.toLocaleString(), ""),
    card("Assessable Income", formatAud(summary.assessableIncome), "income"),
    card("Likely Deductions", formatAud(summary.likelyDeductions), "deduction"),
    card("Possible Deductions", formatAud(summary.possibleDeductions), "deduction"),
    card("Included Deduction Total", formatAud(summary.includedDeductions), "deduction"),
    card("Merchants in FY", summary.merchantCount.toLocaleString(), ""),
    card("Merchants Resolved", `${summary.resolvedMerchants.toLocaleString()}/${summary.merchantCount.toLocaleString()}`, ""),
    card("Internal Transfers Excluded", formatAud(summary.internalTransfers), "review"),
    card("Non-deductible", formatAud(summary.nonDeductible), "review"),
    card("Review Amount", formatAud(summary.reviewAmount), "review")
  ];

  elements.summaryCards.innerHTML = cards.join("");

  elements.summaryMeta.textContent = `FY${fy} (${fyPeriodText(fy)}). Classification is merchant-first; PocketSmith categories are fallback only when merchant text is inconclusive. Merchants resolved with web intelligence: ${summary.resolvedMerchants.toLocaleString()}/${summary.merchantCount.toLocaleString()}. ${
    options.includePossible ? "Possible deductions are included in draft totals." : "Possible deductions are excluded from draft totals."
  }`;

  renderIncomeRows(summary.incomeRows);
  renderDeductionRows(summary.deductionRows);
  renderReviewRows(summary.reviewRows);
  renderMerchantRows(state.merchantGroups, state.merchantIntelByLookup);

  elements.submissionText.value = analysis.markdown;
}

function renderIncomeRows(rows) {
  if (!rows.length) {
    elements.incomeTableBody.innerHTML = "<tr><td colspan=\"4\">No assessable income classified for this FY.</td></tr>";
    return;
  }

  elements.incomeTableBody.innerHTML = rows
    .map(
      (row) => `<tr>
          <td>${escapeHtml(row.category)}</td>
          <td>${row.count}</td>
          <td class="amount">${escapeHtml(formatAud(row.amount))}</td>
          <td>${confidenceChip(dominantConfidence(row))}</td>
        </tr>`
    )
    .join("");
}

function renderDeductionRows(rows) {
  if (!rows.length) {
    elements.deductionTableBody.innerHTML = "<tr><td colspan=\"5\">No deduction candidates classified for this FY.</td></tr>";
    return;
  }

  elements.deductionTableBody.innerHTML = rows
    .map((row) => {
      const confidence = dominantConfidence(row);
      const treatmentLabel = row.treatment === "deduction_likely" ? "Likely" : "Possible";
      return `<tr>
        <td>${escapeHtml(row.atoLabel || "-")}</td>
        <td>${escapeHtml(row.category)} <small>(${treatmentLabel})</small></td>
        <td>${row.count}</td>
        <td class="amount">${escapeHtml(formatAud(row.amount))}</td>
        <td>${confidenceChip(confidence)}</td>
      </tr>`;
    })
    .join("");
}

function renderReviewRows(rows) {
  if (!rows.length) {
    elements.reviewTableBody.innerHTML = "<tr><td colspan=\"5\">No review items were surfaced.</td></tr>";
    return;
  }

  elements.reviewTableBody.innerHTML = rows
    .map(
      (item) => `<tr>
        <td>${escapeHtml(item.tx.dateText)}</td>
        <td>${escapeHtml(item.tx.merchant)}</td>
        <td class="amount">${escapeHtml(formatAud(item.tx.absAmount))}</td>
        <td>${escapeHtml(displayTreatment(item.treatment))}</td>
        <td>${escapeHtml(item.reason)}</td>
      </tr>`
    )
    .join("");
}

function renderMerchantRows(groups, merchantIntelByLookup) {
  if (!groups.length) {
    elements.merchantTableBody.innerHTML = "<tr><td colspan=\"8\">No merchant groups available for this FY.</td></tr>";
    return;
  }

  elements.merchantTableBody.innerHTML = groups
    .map((group) => {
      const intel = merchantIntelByLookup.get(group.lookupKey);
      const sourceLinks = intel && Array.isArray(intel.sourceUrls) && intel.sourceUrls.length
        ? `<span class="source-links">${intel.sourceUrls
            .slice(0, 3)
            .map((url) => `<a href="${escapeHtml(url)}" target="_blank" rel="noreferrer noopener">${escapeHtml(sourceLabel(url))}</a>`)
            .join(" · ")}</span>`
        : "-";

      return `<tr>
        <td class="merchant-cell">${escapeHtml(group.sampleMerchant)}</td>
        <td>${group.transactionCount}</td>
        <td class="amount">${escapeHtml(formatAud(group.totalAbsAmount))}</td>
        <td>${escapeHtml(intel ? intel.businessType : "Not enriched")}</td>
        <td>${escapeHtml(intel && intel.abn ? intel.abn : "-")}</td>
        <td>${escapeHtml(intel && intel.mainPlaceOfBusiness ? intel.mainPlaceOfBusiness : "-")}</td>
        <td>${confidenceChip(normalizeConfidence(intel ? intel.classificationConfidence : "low"))}</td>
        <td>${sourceLinks}</td>
      </tr>`;
    })
    .join("");
}

function sourceLabel(url) {
  try {
    const parsed = new URL(url);
    return parsed.hostname.replace(/^www\./, "");
  } catch (error) {
    return "source";
  }
}

function setMerchantStatus(message, isError = false) {
  elements.merchantStatus.textContent = message;
  elements.merchantStatus.style.color = isError ? "var(--danger)" : "var(--ink-soft)";
}

async function enrichMerchantsForCurrentFy() {
  if (!state.currentFyTransactions.length || !state.selectedFy) {
    setMerchantStatus("Run analysis for an FY before merchant enrichment.", true);
    return;
  }

  for (const group of state.merchantGroups) {
    if (state.merchantIntelByLookup.has(group.lookupKey)) {
      continue;
    }
    if (isSyntheticMerchantLookupKey(group.lookupKey)) {
      state.merchantIntelByLookup.set(group.lookupKey, {
        lookupKey: group.lookupKey,
        merchantRaw: group.sampleMerchant,
        merchantLookupName: group.lookupKey,
        businessType: "Internal transfer / banking event",
        businessCategory: "banking_event",
        classificationConfidence: "high",
        classificationReason: "Derived from canonical merchant key; no external merchant lookup required.",
        abn: "",
        abnName: "",
        abnEntityType: "",
        abnStatus: "",
        mainPlaceOfBusiness: "",
        sourceUrls: [],
        updatedAt: new Date().toISOString()
      });
    }
  }

  const pending = state.merchantGroups.filter(
    (group) => !state.merchantIntelByLookup.has(group.lookupKey) && !isSyntheticMerchantLookupKey(group.lookupKey)
  );
  if (!pending.length) {
    setMerchantStatus("All merchants in this FY are already enriched.");
    elements.downloadMerchantsBtn.disabled = false;
    rebuildAnalysisFromState();
    renderMerchantRows(state.merchantGroups, state.merchantIntelByLookup);
    return;
  }

  elements.enrichMerchantsBtn.disabled = true;
  elements.downloadMerchantsBtn.disabled = true;
  setMerchantStatus(`Starting merchant web enrichment for ${pending.length.toLocaleString()} merchants ...`);

  try {
    let processed = 0;
    for (let i = 0; i < pending.length; i += ENRICH_BATCH_SIZE) {
      const batch = pending.slice(i, i + ENRICH_BATCH_SIZE);
      const { items } = await requestMerchantEnrichment(batch);

      for (const item of items) {
        const key = item.lookupKey || item.merchantLookupKey || deriveMerchantLookupKey(item.merchantRaw || item.merchant || "");
        if (!key) {
          continue;
        }
        state.merchantIntelByLookup.set(key, item);
      }

      processed += batch.length;
      renderMerchantRows(state.merchantGroups, state.merchantIntelByLookup);
      setMerchantStatus(`Enriched ${processed.toLocaleString()} of ${pending.length.toLocaleString()} merchants ...`);
    }

    elements.downloadMerchantsBtn.disabled = false;
    rebuildAnalysisFromState();
    setMerchantStatus(`Merchant enrichment complete. ${state.merchantIntelByLookup.size.toLocaleString()} merchants now classified.`);
  } catch (error) {
    setMerchantStatus(
      `${error.message || "Merchant enrichment failed."} Ensure the app is running via \`npm start\` so /api/enrich-merchants is available.`,
      true
    );
  } finally {
    elements.enrichMerchantsBtn.disabled = false;
  }
}

async function requestMerchantEnrichment(batch) {
  const payload = {
    merchants: batch.map((group) => ({
      lookupKey: group.lookupKey,
      merchant: group.sampleMerchant
    }))
  };

  const response = await fetch("/api/enrich-merchants", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Enrichment API error (${response.status}): ${text || "unknown error"}`);
  }

  const data = await response.json();
  if (!data || !Array.isArray(data.items)) {
    throw new Error("Enrichment API returned an invalid payload.");
  }
  return data;
}

function rebuildAnalysisFromState() {
  if (!state.currentFyTransactions.length || !state.selectedFy || !state.selectedOptions) {
    return;
  }

  const analysis = analyzeTransactions(
    state.currentFyTransactions,
    state.selectedFy,
    state.selectedOptions,
    state.merchantIntelByLookup,
    state.merchantGroups.length
  );

  state.analysis = analysis;
  renderAnalysis(analysis);
}

function downloadMerchantsCsv() {
  if (!state.merchantGroups.length) {
    setMerchantStatus("No merchant groups available to export.", true);
    return;
  }

  const csv = buildMerchantCsv(state.merchantGroups, state.merchantIntelByLookup);
  const fy = state.selectedFy || "unknown";
  const fileName = `merchant-intelligence-FY${fy}.csv`;
  downloadText(fileName, csv, "text/csv;charset=utf-8");
  setMerchantStatus(`Downloaded ${fileName}.`);
}

function buildMerchantCsv(groups, merchantIntelByLookup) {
  const headers = [
    "Merchant Lookup Key",
    "Sample Merchant",
    "Transaction Count",
    "Gross Amount",
    "Business Type",
    "Business Category",
    "Confidence",
    "ABN",
    "ABN Name",
    "ABN Entity Type",
    "Main Place of Business",
    "Sources"
  ];

  const lines = [headers.join(",")];
  for (const group of groups) {
    const intel = merchantIntelByLookup.get(group.lookupKey) || null;
    const row = [
      group.lookupKey,
      group.sampleMerchant,
      group.transactionCount,
      group.totalAbsAmount.toFixed(2),
      intel ? intel.businessType : "",
      intel ? intel.businessCategory : "",
      intel ? normalizeConfidence(intel.classificationConfidence) : "low",
      intel ? intel.abn : "",
      intel ? intel.abnName : "",
      intel ? intel.abnEntityType : "",
      intel ? intel.mainPlaceOfBusiness : "",
      intel && Array.isArray(intel.sourceUrls) ? intel.sourceUrls.join(" | ") : ""
    ];
    lines.push(row.map(csvEscape).join(","));
  }

  return lines.join("\n");
}

function displayTreatment(treatment) {
  switch (treatment) {
    case "income_assessable":
      return "Assessable income";
    case "income_review":
      return "Income review";
    case "deduction_likely":
      return "Likely deduction";
    case "deduction_possible":
      return "Possible deduction";
    case "non_deductible":
      return "Non-deductible";
    case "internal_transfer":
      return "Internal transfer";
    case "excluded_refund":
      return "Refund/reversal excluded";
    default:
      return "Review";
  }
}

function confidenceChip(level) {
  return `<span class="chip ${level}">${level}</span>`;
}

function card(label, value, tone) {
  const className = tone ? `card ${tone}` : "card";
  return `<article class="${className}"><div class="label">${escapeHtml(label)}</div><div class="value">${escapeHtml(value)}</div></article>`;
}

function copySubmission() {
  if (!state.analysis) {
    setStatus("Run analysis first, then copy the submission.", true);
    return;
  }

  navigator.clipboard
    .writeText(elements.submissionText.value)
    .then(() => setStatus("Submission markdown copied to clipboard."))
    .catch(() => setStatus("Clipboard copy failed. You can still download the markdown file.", true));
}

function downloadSubmission() {
  if (!state.analysis) {
    setStatus("Run analysis first, then download submission markdown.", true);
    return;
  }

  const fileName = `tax-submission-FY${state.analysis.fy}.md`;
  downloadText(fileName, state.analysis.markdown, "text/markdown;charset=utf-8");
  setStatus(`Downloaded ${fileName}.`);
}

function downloadClassifiedCsv() {
  if (!state.analysis) {
    setStatus("Run analysis first, then download classified CSV.", true);
    return;
  }

  const csv = buildClassifiedCsv(state.analysis.classified);
  const fileName = `classified-transactions-FY${state.analysis.fy}.csv`;
  downloadText(fileName, csv, "text/csv;charset=utf-8");
  setStatus(`Downloaded ${fileName}.`);
}

function buildClassifiedCsv(classified) {
  const headers = [
    "Date",
    "Merchant",
    "Merchant Lookup Key",
    "Account",
    "Amount",
    "Currency",
    "Treatment",
    "Tax Category",
    "ATO Label",
    "Confidence",
    "Merchant Business Type",
    "Merchant Business Category",
    "Merchant ABN",
    "Merchant Main Place",
    "Reason",
    "Labels",
    "PocketSmith Category",
    "Parent Category",
    "ID"
  ];

  const lines = [headers.join(",")];
  for (const item of classified) {
    const intel = state.merchantIntelByLookup.get(item.tx.merchantLookupKey) || null;
    const row = [
      item.tx.dateText,
      item.tx.merchant,
      item.tx.merchantLookupKey,
      item.tx.account,
      item.tx.amount.toFixed(2),
      item.tx.currency,
      displayTreatment(item.treatment),
      item.taxCategory,
      item.atoLabel,
      item.confidence,
      intel ? intel.businessType : "",
      intel ? intel.businessCategory : "",
      intel ? intel.abn : "",
      intel ? intel.mainPlaceOfBusiness : "",
      item.reason,
      item.tx.labels.join("|"),
      item.tx.category,
      item.tx.parentCategory,
      item.tx.pocketId
    ];
    lines.push(row.map(csvEscape).join(","));
  }

  return lines.join("\n");
}

function csvEscape(value) {
  const raw = String(value == null ? "" : value);
  if (!/[",\n]/.test(raw)) {
    return raw;
  }
  return `"${raw.replace(/"/g, '""')}"`;
}

function downloadText(filename, text, type) {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function clearResults() {
  elements.summaryPanel.classList.add("hidden");
  elements.incomePanel.classList.add("hidden");
  elements.deductionPanel.classList.add("hidden");
  elements.reviewPanel.classList.add("hidden");
  elements.merchantPanel.classList.add("hidden");
  elements.submissionPanel.classList.add("hidden");
  elements.summaryCards.innerHTML = "";
  elements.summaryMeta.textContent = "";
  elements.incomeTableBody.innerHTML = "";
  elements.deductionTableBody.innerHTML = "";
  elements.reviewTableBody.innerHTML = "";
  elements.merchantTableBody.innerHTML = "";
  elements.enrichMerchantsBtn.disabled = true;
  elements.downloadMerchantsBtn.disabled = true;
  setMerchantStatus("Merchant enrichment has not been run.");
  elements.submissionText.value = "";
}

function setStatus(message, isError = false) {
  elements.analysisStatus.textContent = message;
  elements.analysisStatus.style.color = isError ? "var(--danger)" : "var(--ink-soft)";
}

function formatAud(value) {
  return new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency: CURRENCY,
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(Number(value) || 0);
}

function formatDateTimeSydney(date) {
  return new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Sydney",
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeMarkdown(value) {
  return String(value == null ? "" : value)
    .replace(/\|/g, "\\|")
    .replace(/\n/g, " ")
    .trim();
}
