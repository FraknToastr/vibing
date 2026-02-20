const MAX_ROWS = 500;
const DEFAULT_NEW_HEADERS = [
  "Asset Class",
  "Asset SubClass",
  "Asset Type",
  "Asset SubType",
  "Component Name",
  "Asset ID (if known)",
  "Quantity",
  "Unit of Measurement",
  "Unit Cost",
  "Total Cost",
  "Component Type",
  "Asset Category",
  "Financial Class",
  "Financial SubClass",
  "Asset Network Measure Type",
  "Allocate Project-Wide Costs to this Asset?",
  "Capitalise This Asset? ",
  "Useful Life",
  "Valuation Record ID",
  "Asset Name",
  "Primary Material",
  "Make/ Model",
  "Location",
  "Valuation Date",
  "Revaluation Date Built",
  "Valuation Record Type"
];
const DEFAULT_RENEWED_HEADERS = [
  "Asset Class",
  "Asset SubClass",
  "Asset Type",
  "Asset SubType",
  "Component Name",
  "Asset ID",
  "Quantity",
  "Unit of Measurement",
  "Unit Cost",
  "Total Cost",
  "Component Type",
  "Upgrade (%)",
  "Upgrade $",
  "Renewal $",
  "Allocate Project-Wide Costs to this Asset?",
  "Capitalise This Asset? ",
  "Treatment Type",
  "% of Asset Renewed",
  "Useful Life",
  "Condition Rating",
  "Valuation Date",
  "Date Built",
  "Valuation Record ID",
  "Valuation Component Name"
];
const DEFAULT_DISPOSED_HEADERS = [
  "Asset Class",
  "Asset SubClass",
  "Asset Type",
  "Asset SubType",
  "Component Name",
  "Asset ID",
  "Asset Name Or Description",
  "Location",
  "Disposal Date",
  "Reason",
  "Disposal Type",
  "Valuation Record ID/ FAR ID",
  "Valuation Record Type",
  "Valuation Date",
  "Valuation Component Name"
];
const DEFAULT_SERVICE_HEADERS = [
  "Asset ID",
  "Asset SubType",
  "Component Name",
  "Component Type",
  "Service Criteria Type",
  "Assessment Date",
  "Score",
  "Assessed By Resource Name"
];

const LOOKUP_FILES = {
  assetLookup: "Asset Lookup.csv",
  componentLookup: "Component Lookup.csv",
  componentFinancial: "Component Financial Class and Subclass Lookup.csv",
  valuationComponent: "Valuation Component Lookup.csv",
  serviceCriteria: "Service Criteria Lookup.csv"
};

const HEADER_FILES = {
  new: "PAR New Asset Headers.csv",
  renewed: "PAR Renewal Headers.csv",
  disposed: "PAR Disposal Headers.csv",
  serviceCriteria: "PAR Service Criteria Headers.csv"
};

const REMOTE_BASE = "https://raw.githubusercontent.com/FraknToastr/PAR/main/";

const lookupStatus = document.getElementById("lookupStatus");
const toggleWidthBtn = document.getElementById("toggleWidth");
const resetAllBtn = document.getElementById("resetAll");
const saveSessionBtn = document.getElementById("saveSession");
const loadSessionBtn = document.getElementById("loadSession");
const sessionFileInput = document.getElementById("sessionFileInput");
const projectCodeInput = document.getElementById("projectCodeInput");
const applyProjectCodeBtn = document.getElementById("applyProjectCode");
const loadLookupsBtn = document.getElementById("loadLookups");
const loadLookupsLocalBtn = document.getElementById("loadLookupsLocal");
const lookupBulkInput = document.getElementById("lookupBulkInput");
const lookupSections = Array.from(document.querySelectorAll(".lookup-item"));

const newAssetsTable = document.getElementById("newAssetsTable");
const addNewRowBtn = document.getElementById("addNewRow");
const duplicateNewRowBtn = document.getElementById("duplicateNewRow");
const duplicateCountSelect = document.getElementById("duplicateCount");
const insertNewRowBtn = document.getElementById("insertNewRow");
const insertDuplicateRowBtn = document.getElementById("insertDuplicateRow");
const removeSelectedRowsBtn = document.getElementById("removeSelectedRows");
const exportNewBtn = document.getElementById("exportNewCsv");
const newAssetStatus = document.getElementById("newAssetStatus");
const newAssetResult = document.getElementById("newAssetResult");
const resetNewBtn = document.getElementById("resetNew");

const renewedAssetsTable = document.getElementById("renewedAssetsTable");
const renewedStatus = document.getElementById("renewedStatus");
const renewedLoadIdsBtn = document.getElementById("renewedLoadIds");
const renewedIdFileInput = document.getElementById("renewedIdFile");
const renewedApplyIdsBtn = document.getElementById("renewedApplyIds");
const renewedTogglePasteBtn = document.getElementById("renewedTogglePaste");
const renewedValidateBtn = document.getElementById("renewedValidate");
const renewedToggleFieldsBtn = document.getElementById("renewedToggleFields");
const exportRenewedBtn = document.getElementById("exportRenewedCsv");
const renewedIdText = document.getElementById("renewedIdText");
const resetRenewedBtn = document.getElementById("resetRenewed");
const disposedAssetsTable = document.getElementById("disposedAssetsTable");
const disposedStatus = document.getElementById("disposedStatus");
const disposedLoadIdsBtn = document.getElementById("disposedLoadIds");
const disposedIdFileInput = document.getElementById("disposedIdFile");
const disposedApplyIdsBtn = document.getElementById("disposedApplyIds");
const disposedTogglePasteBtn = document.getElementById("disposedTogglePaste");
const disposedValidateBtn = document.getElementById("disposedValidate");
const disposedToggleFieldsBtn = document.getElementById("disposedToggleFields");
const exportDisposedBtn = document.getElementById("exportDisposedCsv");
const disposedIdText = document.getElementById("disposedIdText");
const resetDisposedBtn = document.getElementById("resetDisposed");
const serviceCriteriaTable = document.getElementById("serviceCriteriaTable");
const addServiceCriteriaRowBtn = document.getElementById("addServiceCriteriaRow");
const duplicateServiceCriteriaRowBtn = document.getElementById("duplicateServiceCriteriaRow");
const serviceCriteriaDuplicateCountSelect = document.getElementById("serviceCriteriaDuplicateCount");
const insertServiceCriteriaRowBtn = document.getElementById("insertServiceCriteriaRow");
const insertDuplicateServiceCriteriaRowBtn = document.getElementById("insertDuplicateServiceCriteriaRow");
const removeServiceCriteriaRowsBtn = document.getElementById("removeServiceCriteriaRows");
const validateServiceCriteriaBtn = document.getElementById("validateServiceCriteria");
const exportServiceCriteriaBtn = document.getElementById("exportServiceCriteriaCsv");
const serviceCriteriaStatus = document.getElementById("serviceCriteriaStatus");
const resetServiceCriteriaBtn = document.getElementById("resetServiceCriteria");
const tabButtons = Array.from(document.querySelectorAll(".tab-btn"));
const tabPanels = Array.from(document.querySelectorAll(".tab-panel"));

const lookupTables = {};
let newHeaders = [...DEFAULT_NEW_HEADERS];
let newRows = [];
let subtypeOptions = [];
let renewedHeaders = [...DEFAULT_RENEWED_HEADERS];
let renewedRows = [];
let disposedHeaders = [...DEFAULT_DISPOSED_HEADERS];
let disposedRows = [];
let serviceCriteriaHeaders = [...DEFAULT_SERVICE_HEADERS];
let serviceCriteriaRows = [];
const selectedRowIndices = new Set();
const selectedRenewedRowIndices = new Set();
const selectedDisposedRowIndices = new Set();
const selectedServiceCriteriaRowIndices = new Set();
const invalidRenewedCells = new Set();
const invalidDisposedCells = new Set();
const warnRenewedCells = new Set();
const warnDisposedCells = new Set();
const validationStateRenewed = new Map();
const validationStateDisposed = new Map();
let hideRenewedNonCore = false;
let hideDisposedNonCore = false;
let projectCodeValue = "";
const RENEWED_HIDE_OVERRIDES = new Set([
  normalizeHeader("Allocate Project-Wide Costs to this Asset?"),
  normalizeHeader("Capitalise This Asset?"),
  normalizeHeader("% of Asset Renewed"),
  normalizeHeader("Valuation Date"),
  normalizeHeader("Valuation Record ID")
]);
const DISPOSED_HIDE_OVERRIDES = new Set([
  normalizeHeader("Asset Name Or Description"),
  normalizeHeader("Valuation Record ID/ FAR ID"),
  normalizeHeader("Valuation Record Type"),
  normalizeHeader("Valuation Date")
]);
const SERVICE_CRITERIA_SELECT_FIELDS = new Set([
  normalizeHeader("Asset SubType"),
  normalizeHeader("Component Name"),
  normalizeHeader("Component Type"),
  normalizeHeader("Service Criteria Type")
]);
const FORCE_SELECT_FIELDS = new Set([
  "Asset Category",
  "Asset Class",
  "Asset SubClass",
  "Asset Type",
  "Asset SubType"
]);
const NEW_ROW_DEFAULTS = {
  "Valuation Record Type": "Constructed"
};
const RESET_CASCADE = {
  "Asset Category": [
    "Asset Class",
    "Asset SubClass",
    "Asset Type",
    "Asset SubType",
    "Component Name",
    "Component Type",
    "Financial Class",
    "Financial SubClass"
  ],
  "Asset Class": [
    "Asset SubClass",
    "Asset Type",
    "Asset SubType",
    "Component Name",
    "Component Type",
    "Financial Class",
    "Financial SubClass"
  ],
  "Asset SubClass": [
    "Asset Type",
    "Asset SubType",
    "Component Name",
    "Component Type",
    "Financial Class",
    "Financial SubClass"
  ],
  "Asset Type": [
    "Asset SubType",
    "Component Name",
    "Component Type",
    "Financial Class",
    "Financial SubClass"
  ],
  "Asset SubType": [
    "Component Name",
    "Component Type",
    "Financial Class",
    "Financial SubClass"
  ],
  "Component Name": ["Component Type", "Financial Class", "Financial SubClass"],
  "Component Type": ["Financial Class", "Financial SubClass"],
  "Financial Class": ["Financial SubClass"]
};
const LOOKUP_FIELDS = [
  "Asset Category",
  "Asset Class",
  "Asset SubClass",
  "Asset Type",
  "Asset SubType",
  "Component Name",
  "Component Type",
  "Financial Class",
  "Financial SubClass"
];

function isLookupComplete(rowData, headerMap) {
  const fields = LOOKUP_FIELDS.map((label) => headerMap.get(normalizeHeader(label))).filter(Boolean);
  if (!fields.length) {
    return false;
  }
  return fields.every((field) => String(rowData[field] || "").trim());
}

function parseCSV(text) {
  const cleaned = text.replace(/^\ufeff/, "");
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let i = 0; i < cleaned.length; i += 1) {
    const char = cleaned[i];
    const next = cleaned[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        value += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(value);
      value = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      continue;
    }

    value += char;
  }

  row.push(value);
  rows.push(row);

  return rows;
}

function toCSV(headers, rows) {
  const all = [headers, ...rows];
  return all
    .map((row) => row.map(escapeCSVValue).join(","))
    .join("\n");
}

function escapeCSVValue(value) {
  const text = value == null ? "" : String(value);
  if (text.includes("\n") || text.includes(",") || text.includes('"')) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}

function normalizeHeader(text) {
  return String(text || "")
    .trim()
    .toLowerCase()
    .replace(/[\s_]+/g, "");
}

function normalizeValue(value) {
  return String(value || "").trim().toLowerCase();
}

function sanitizeFilePart(value) {
  return String(value || "")
    .trim()
    .replace(/[^a-z0-9_-]+/gi, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();
}

function buildServiceCriteriaKey(assetId, subtype, compName, compType) {
  const parts = [assetId, subtype, compName, compType].map((value) => normalizeValue(value));
  if (!parts.some((value) => value)) {
    return "";
  }
  return parts.join("||");
}

function buildHeaderMap(headers) {
  const map = new Map();
  headers.forEach((header, index) => {
    map.set(normalizeHeader(header), index);
  });
  return map;
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

function normalizeFileName(name) {
  return String(name || "")
    .trim()
    .toLowerCase();
}

async function fetchTextWithFallback(file) {
  const candidates = [];
  if (file) {
    candidates.push(file);
    if (!/^https?:/i.test(file)) {
      candidates.push(new URL(file, REMOTE_BASE).toString());
    }
  }

  let lastError = null;
  for (const url of candidates) {
    try {
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`Fetch failed: ${response.status}`);
      }
      return await response.text();
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("Fetch failed");
}

function updateLookupStatus(message, isError = false) {
  lookupStatus.textContent = message;
  lookupStatus.style.color = isError ? "#b63a26" : "";
}

function ensureColumns(headers, rows, columns) {
  const headerMap = buildHeaderMap(headers);
  columns.forEach((col) => {
    if (!headerMap.has(normalizeHeader(col))) {
      headers.push(col);
      rows.forEach((row) => row.push(""));
    }
  });
}

function parseLookupTable(text) {
  const parsed = parseCSV(text.trim());
  if (parsed.length === 0 || parsed[0].length === 0) {
    return null;
  }
  const headers = parsed[0];
  const rows = parsed.slice(1);
  const objects = rows.map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[normalizeHeader(header)] = row[index] || "";
    });
    return obj;
  });
  return { headers, rows, objects };
}

function buildLookupFromRows(headers, rows) {
  const objects = rows.map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[normalizeHeader(header)] = row[index] || "";
    });
    return obj;
  });
  return { headers, rows, objects };
}

function buildIndex(objects, keyHeader) {
  const keyNorm = normalizeHeader(keyHeader);
  const map = new Map();
  objects.forEach((obj) => {
    const key = normalizeValue(obj[keyNorm]);
    if (!key) {
      return;
    }
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(obj);
  });
  return map;
}

function buildCompositeIndex(objects, headerA, headerB) {
  const keyA = normalizeHeader(headerA);
  const keyB = normalizeHeader(headerB);
  const map = new Map();
  objects.forEach((obj) => {
    const a = normalizeValue(obj[keyA]);
    const b = normalizeValue(obj[keyB]);
    if (!a || !b) {
      return;
    }
    const key = `${a}||${b}`;
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(obj);
  });
  return map;
}

function ensureValuationIndex() {
  const valuationLookup = lookupTables.valuationComponent;
  if (!valuationLookup?.objects?.length) {
    return false;
  }
  if (!valuationLookup.indexByAssetId || !valuationLookup.indexByAssetId.size) {
    let indexByAssetId = buildIndex(valuationLookup.objects, "Asset_Id");
    if (!indexByAssetId.size) {
      indexByAssetId = buildIndex(valuationLookup.objects, "Asset ID");
    }
    if (!indexByAssetId.size) {
      const sample = valuationLookup.objects[0] || {};
      const key = Object.keys(sample).find((k) => k.includes("assetid"));
      if (key) {
        indexByAssetId = buildIndex(valuationLookup.objects, key);
      }
    }
    valuationLookup.indexByAssetId = indexByAssetId;
  }
  return Boolean(valuationLookup.indexByAssetId?.size);
}

function setValueIfEmpty(rowData, field, value, resolvedFields) {
  if (!field || !value) {
    return false;
  }
  if (String(rowData[field] || "").trim()) {
    return false;
  }
  rowData[field] = value;
  resolvedFields.add(field);
  return true;
}

function loadLookupFromText(id, text, statusEl, displayEl) {
  const trimmed = text.trim();
  if (!trimmed) {
    statusEl.textContent = "Empty.";
    statusEl.style.color = "#b63a26";
    lookupTables[id] = null;
    return;
  }

  const parsed = parseLookupTable(trimmed);
  if (!parsed) {
    statusEl.textContent = "Invalid CSV.";
    statusEl.style.color = "#b63a26";
    lookupTables[id] = null;
    return;
  }

  lookupTables[id] = parsed;
  statusEl.textContent = `Loaded ${parsed.rows.length} rows. ✅`;
  statusEl.style.color = "#1fcf6e";
  if (displayEl) {
    displayEl.value = text;
    displayEl.readOnly = false;
  }

  if (id === "assetLookup") {
    lookupTables.assetLookup.index = buildIndex(parsed.objects, "Asset_Subtype");
    subtypeOptions = Array.from(lookupTables.assetLookup.index.keys())
      .map((value) => parsed.objects.find((obj) => normalizeValue(obj.assetsubtype) === value)?.assetsubtype || value)
      .sort((a, b) => a.localeCompare(b));
  }

  if (id === "componentLookup") {
    lookupTables.componentLookup.index = buildIndex(parsed.objects, "Asset_Subtype");
  }

  if (id === "componentFinancial") {
    lookupTables.componentFinancial.index = buildCompositeIndex(
      parsed.objects,
      "Component_Name",
      "Component_Type"
    );
  }

  if (id === "valuationComponent") {
    lookupTables.valuationComponent.index = buildCompositeIndex(
      parsed.objects,
      "Component_Name",
      "Component_Type"
    );
    ensureValuationIndex();
  }

  updateAllRows();
}

function hydrateLookup(id, headers, rows) {
  const parsed = buildLookupFromRows(headers, rows);
  lookupTables[id] = parsed;
  if (id === "assetLookup") {
    lookupTables.assetLookup.index = buildIndex(parsed.objects, "Asset_Subtype");
    subtypeOptions = Array.from(lookupTables.assetLookup.index.keys())
      .map((value) => parsed.objects.find((obj) => normalizeValue(obj.assetsubtype) === value)?.assetsubtype || value)
      .sort((a, b) => a.localeCompare(b));
  }

  if (id === "componentLookup") {
    lookupTables.componentLookup.index = buildIndex(parsed.objects, "Asset_Subtype");
  }

  if (id === "componentFinancial") {
    lookupTables.componentFinancial.index = buildCompositeIndex(
      parsed.objects,
      "Component_Name",
      "Component_Type"
    );
  }

  if (id === "valuationComponent") {
    lookupTables.valuationComponent.index = buildCompositeIndex(
      parsed.objects,
      "Component_Name",
      "Component_Type"
    );
    ensureValuationIndex();
  }
}

lookupSections.forEach((section) => {
  const id = section.dataset.lookup;
  const fileInput = section.querySelector(".lookupFile");
  const textArea = section.querySelector(".lookupText");
  const statusEl = section.querySelector(".lookupStatus");

  fileInput.addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (!file) {
      return;
    }
    const text = await readFileAsText(file);
    loadLookupFromText(id, text, statusEl, textArea);
  });

  section._loadLookup = () => {
    if (!textArea) {
      return;
    }
    loadLookupFromText(id, textArea.value, statusEl, textArea);
  };
});

function loadLookupsFromRepo() {
  Promise.all(
    lookupSections.map(async (section) => {
      const id = section.dataset.lookup;
      const statusEl = section.querySelector(".lookupStatus");
      const textArea = section.querySelector(".lookupText");
      const repoFile = LOOKUP_FILES[id];
      if (!repoFile) {
        section._loadLookup();
        return;
      }
      try {
        const text = await fetchTextWithFallback(repoFile);
        loadLookupFromText(id, text, statusEl, textArea);
      } catch (error) {
        section._loadLookup();
      }
    })
  ).then(() => {
    const loadedCount = Object.values(lookupTables).filter(Boolean).length;
    updateLookupStatus(`Loaded ${loadedCount} lookup tables.`);
  });
}

loadLookupsBtn.addEventListener("click", loadLookupsFromRepo);

function buildBulkFileMap(fileList) {
  const map = new Map();
  Array.from(fileList || []).forEach((file) => {
    map.set(normalizeFileName(file.name), file);
  });
  return map;
}

async function handleBulkLookupFiles(event) {
  const files = event.target.files;
  if (!files || !files.length) {
    return;
  }

  const allowedNames = new Set(
    Object.values(LOOKUP_FILES).map((name) => normalizeFileName(name))
  );
  const filteredFiles = Array.from(files).filter((file) =>
    allowedNames.has(normalizeFileName(file.name))
  );
  if (!filteredFiles.length) {
    updateLookupStatus("No valid lookup CSVs selected.", true);
    return;
  }
  const fileMap = buildBulkFileMap(filteredFiles);
  const lookupPromises = lookupSections.map(async (section) => {
    const id = section.dataset.lookup;
    const statusEl = section.querySelector(".lookupStatus");
    const textArea = section.querySelector(".lookupText");
    const repoFile = LOOKUP_FILES[id];
    const file = fileMap.get(normalizeFileName(repoFile));
    if (!file) {
      return;
    }
    const text = await readFileAsText(file);
    loadLookupFromText(id, text, statusEl, textArea);
  });

  await Promise.all(lookupPromises);
  const loadedCount = Object.values(lookupTables).filter(Boolean).length;
  updateLookupStatus(`Loaded ${loadedCount} lookup tables.`);
  updateAllRows();
}

if (loadLookupsLocalBtn && lookupBulkInput) {
  loadLookupsLocalBtn.addEventListener("click", () => {
    lookupBulkInput.value = "";
    lookupBulkInput.click();
  });
  lookupBulkInput.addEventListener("change", handleBulkLookupFiles);
}

function buildNewAssetHeaders() {
  const headerMap = new Map();
  newHeaders.forEach((header) => headerMap.set(normalizeHeader(header), header));
  return headerMap;
}

function buildRenewedHeaders() {
  const headerMap = new Map();
  renewedHeaders.forEach((header) => headerMap.set(normalizeHeader(header), header));
  return headerMap;
}

function buildDisposedHeaders() {
  const headerMap = new Map();
  disposedHeaders.forEach((header) => headerMap.set(normalizeHeader(header), header));
  return headerMap;
}

function buildServiceCriteriaHeaders() {
  const headerMap = new Map();
  serviceCriteriaHeaders.forEach((header) => headerMap.set(normalizeHeader(header), header));
  return headerMap;
}

function isCoreField(header) {
  const normalized = normalizeHeader(header);
  if (!normalized) {
    return false;
  }
  return (
    normalized.includes("asset") ||
    normalized.includes("component") ||
    normalized.includes("valuation")
  );
}

function isDateField(header) {
  const normalized = normalizeHeader(header);
  return normalized.includes("date");
}

function resolveNewAssetRow(rowData, headerMap) {
  const options = {};
  const resolvedFields = new Set();
  const ambiguousFields = new Set();

  const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
  const typeField = headerMap.get(normalizeHeader("Asset Type"));
  const subClassField = headerMap.get(normalizeHeader("Asset SubClass"));
  const classField = headerMap.get(normalizeHeader("Asset Class"));
  const categoryField = headerMap.get(normalizeHeader("Asset Category"));
  const compNameField = headerMap.get(normalizeHeader("Component Name"));
  const compTypeField = headerMap.get(normalizeHeader("Component Type"));
  const finClassField = headerMap.get(normalizeHeader("Financial Class"));
  const finSubClassField = headerMap.get(normalizeHeader("Financial SubClass"));

  const assetLookup = lookupTables.assetLookup;
  if (assetLookup?.index) {
    const currentSubtype = normalizeValue(rowData[subtypeField]);
    const currentType = normalizeValue(rowData[typeField]);
    const currentSubClass = normalizeValue(rowData[subClassField]);
    const currentClass = normalizeValue(rowData[classField]);
    const currentCategory = normalizeValue(rowData[categoryField]);

    let candidates = assetLookup.objects;
    if (currentSubtype) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetsubtype) === currentSubtype);
    }
    if (currentType) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assettype) === currentType);
    }
    if (currentSubClass) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetsubclass) === currentSubClass);
    }
    if (currentClass) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetclass) === currentClass);
    }
    if (currentCategory) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetcategory) === currentCategory);
    }

    const candidateValues = {
      [categoryField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assetcategory).filter(Boolean))
      ).sort((a, b) => a.localeCompare(b)),
      [classField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assetclass).filter(Boolean))
      ).sort((a, b) => a.localeCompare(b)),
      [subClassField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assetsubclass).filter(Boolean))
      ).sort((a, b) => a.localeCompare(b)),
      [typeField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assettype).filter(Boolean))
      ).sort((a, b) => a.localeCompare(b)),
      [subtypeField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assetsubtype).filter(Boolean))
      ).sort((a, b) => a.localeCompare(b))
    };

    [categoryField, classField, subClassField, typeField, subtypeField].forEach((field) => {
      if (!field) {
        return;
      }
      if (candidateValues[field]?.length) {
        options[field] = candidateValues[field];
        if (!String(rowData[field] || "").trim()) {
          if (candidateValues[field].length > 1) {
            ambiguousFields.add(field);
          } else if (candidateValues[field].length === 1) {
            setValueIfEmpty(rowData, field, candidateValues[field][0], resolvedFields);
          }
        }
      }
    });

    if (candidates.length === 1) {
      const match = candidates[0];
      setValueIfEmpty(rowData, categoryField, match.assetcategory, resolvedFields);
      setValueIfEmpty(rowData, classField, match.assetclass, resolvedFields);
      setValueIfEmpty(rowData, subClassField, match.assetsubclass, resolvedFields);
      setValueIfEmpty(rowData, typeField, match.assettype, resolvedFields);
      setValueIfEmpty(rowData, subtypeField, match.assetsubtype, resolvedFields);
    }
  }

  const componentLookup = lookupTables.componentLookup;
  const subtypeValue = normalizeValue(rowData[subtypeField]);
  if (componentLookup?.index && subtypeValue) {
    const compCandidates = componentLookup.index.get(subtypeValue) || [];
    if (compCandidates.length === 1) {
      const match = compCandidates[0];
      setValueIfEmpty(rowData, compNameField, match.componentname, resolvedFields);
      setValueIfEmpty(rowData, compTypeField, match.componenttype, resolvedFields);
    } else if (compCandidates.length > 1) {
      const currentName = normalizeValue(rowData[compNameField]);
      const currentType = normalizeValue(rowData[compTypeField]);
      let filtered = compCandidates;
      if (currentName) {
        filtered = filtered.filter((candidate) => normalizeValue(candidate.componentname) === currentName);
      }
      if (currentType) {
        filtered = filtered.filter((candidate) => normalizeValue(candidate.componenttype) === currentType);
      }

      if (filtered.length === 1) {
        const match = filtered[0];
        setValueIfEmpty(rowData, compNameField, match.componentname, resolvedFields);
        setValueIfEmpty(rowData, compTypeField, match.componenttype, resolvedFields);
      } else {
        if (!currentName) {
          const names = Array.from(
            new Set(compCandidates.map((candidate) => candidate.componentname).filter(Boolean))
          ).sort((a, b) => a.localeCompare(b));
          if (names.length > 1) {
            options[compNameField] = names;
            ambiguousFields.add(compNameField);
          } else if (names.length === 1) {
            setValueIfEmpty(rowData, compNameField, names[0], resolvedFields);
          }
        }
        if (currentName && !currentType) {
          const types = Array.from(
            new Set(
              compCandidates
                .filter((candidate) => normalizeValue(candidate.componentname) === currentName)
                .map((candidate) => candidate.componenttype)
                .filter(Boolean)
            )
          ).sort((a, b) => a.localeCompare(b));
          if (types.length > 1) {
            options[compTypeField] = types;
            ambiguousFields.add(compTypeField);
          } else if (types.length === 1) {
            setValueIfEmpty(rowData, compTypeField, types[0], resolvedFields);
          }
        }
      }
    }
  }

  const financialLookup = lookupTables.componentFinancial;
  if (financialLookup?.index) {
    const compName = normalizeValue(rowData[compNameField]);
    const compType = normalizeValue(rowData[compTypeField]);
    if (compName && compType) {
      const key = `${compName}||${compType}`;
      const matches = financialLookup.index.get(key) || [];
      if (matches.length === 1) {
        const match = matches[0];
        setValueIfEmpty(rowData, finClassField, match.componentfinancialclass, resolvedFields);
        setValueIfEmpty(rowData, finSubClassField, match.componentfinancialsubclass, resolvedFields);
      } else if (matches.length > 1) {
        if (!rowData[finClassField]) {
          const classes = Array.from(
            new Set(matches.map((candidate) => candidate.componentfinancialclass).filter(Boolean))
          );
          if (classes.length > 1) {
            options[finClassField] = classes;
            ambiguousFields.add(finClassField);
          } else if (classes.length === 1) {
            setValueIfEmpty(rowData, finClassField, classes[0], resolvedFields);
          }
        }
        if (!rowData[finSubClassField]) {
          const subclasses = Array.from(
            new Set(matches.map((candidate) => candidate.componentfinancialsubclass).filter(Boolean))
          );
          if (subclasses.length > 1) {
            options[finSubClassField] = subclasses;
            ambiguousFields.add(finSubClassField);
          } else if (subclasses.length === 1) {
            setValueIfEmpty(rowData, finSubClassField, subclasses[0], resolvedFields);
          }
        }
      }
    }
  }

  return { options, resolvedFields, ambiguousFields };
}

function getResetTargets(field, headerMap) {
  const normalized = normalizeHeader(field);
  const matchingLabel = Object.keys(RESET_CASCADE).find(
    (label) => normalizeHeader(label) === normalized
  );
  const targets = new Set();
  if (field) {
    targets.add(field);
  }
  if (matchingLabel) {
    RESET_CASCADE[matchingLabel].forEach((label) => {
      const actual = headerMap.get(normalizeHeader(label));
      if (actual) {
        targets.add(actual);
      }
    });
  }
  return Array.from(targets);
}

function handleFieldReset(event) {
  event.preventDefault();
  event.stopPropagation();
  const field = event.currentTarget.dataset.field;
  const rowIndex = Number(event.currentTarget.dataset.row);
  if (!newRows[rowIndex]) {
    return;
  }
  const headerMap = buildNewAssetHeaders();
  const targets = getResetTargets(field, headerMap);
  targets.forEach((target) => {
    newRows[rowIndex][target] = NEW_ROW_DEFAULTS[target] || "";
  });
  renderNewAssetsTable();
}

function handleCellInput(event) {
  const field = event.target.dataset.field;
  const rowIndex = Number(event.target.dataset.row);
  if (!newRows[rowIndex]) {
    return;
  }
  newRows[rowIndex][field] = event.target.value;
  syncServiceCriteriaFromSources();
}

function createCellInput(
  field,
  value,
  options,
  rowIndex,
  colIndex,
  headerMap,
  resolvedFields,
  ambiguousFields,
  forceSelect,
  addFillDown = false
) {
  const td = document.createElement("td");
  const isAmbiguous = ambiguousFields.has(field);
  const isResolved = resolvedFields.has(field);

  if (isAmbiguous) {
    td.classList.add("cell-ambiguous");
  } else if (isResolved) {
    td.classList.add("cell-resolved");
  }

  const useSelect = forceSelect || (options && options.length);
  const controlWrap = document.createElement("div");
  controlWrap.className = "cell-control";

  if (useSelect) {
    const select = document.createElement("select");
    select.className = "cell-select";
    select.dataset.field = field;
    select.dataset.row = rowIndex;
    select.dataset.col = colIndex;

    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "";
    select.appendChild(blank);

    const optionValues = Array.isArray(options) ? [...options] : [];
    if (value && !optionValues.includes(value)) {
      optionValues.unshift(value);
    }
    optionValues.forEach((optionValue) => {
      const option = document.createElement("option");
      option.value = optionValue;
      option.textContent = optionValue;
      select.appendChild(option);
    });

    select.value = value || "";
    select.addEventListener("change", handleCellChange);
    controlWrap.appendChild(select);
  } else {
    const input = document.createElement("input");
    input.type = isDateField(field) ? "date" : "text";
    input.className = "cell-input";
    input.value = value || "";
    input.dataset.field = field;
    input.dataset.row = rowIndex;
    input.dataset.col = colIndex;
    input.addEventListener("input", handleCellInput);
    controlWrap.appendChild(input);
  }

  if (addFillDown) {
    const fillBtn = document.createElement("button");
    fillBtn.type = "button";
    fillBtn.className = "field-filldown";
    fillBtn.textContent = "⇩";
    fillBtn.title = `Fill down ${field}`;
    fillBtn.setAttribute("aria-label", `Fill down ${field}`);
    fillBtn.addEventListener("click", () => {
      fillDownNewDateField(field);
    });
    controlWrap.appendChild(fillBtn);
  }

  const resetBtn = document.createElement("button");
  resetBtn.type = "button";
  resetBtn.className = "field-reset";
  resetBtn.textContent = "🧽";
  resetBtn.title = `Reset ${field}`;
  resetBtn.setAttribute("aria-label", `Reset ${field}`);
  resetBtn.dataset.field = field;
  resetBtn.dataset.row = rowIndex;
  resetBtn.dataset.col = colIndex;
  resetBtn.addEventListener("click", handleFieldReset);
  controlWrap.appendChild(resetBtn);

  td.appendChild(controlWrap);
  return td;
}

function handleRenewedCellInput(event) {
  const field = event.target.dataset.field;
  const rowIndex = Number(event.target.dataset.row);
  if (!renewedRows[rowIndex]) {
    return;
  }
  renewedRows[rowIndex][field] = event.target.value;
  invalidRenewedCells.delete(`${rowIndex}||${field}`);
  warnRenewedCells.delete(`${rowIndex}||${field}`);
  validationStateRenewed.delete(rowIndex);
  syncServiceCriteriaFromSources();
}

function handleRenewedCellChange(event) {
  const field = event.target.dataset.field;
  const rowIndex = Number(event.target.dataset.row);
  if (!renewedRows[rowIndex]) {
    return;
  }
  renewedRows[rowIndex][field] = event.target.value;
  invalidRenewedCells.delete(`${rowIndex}||${field}`);
  warnRenewedCells.delete(`${rowIndex}||${field}`);
  validationStateRenewed.delete(rowIndex);
  renderRenewedAssetsTable();
  syncServiceCriteriaFromSources();
}

function handleDisposedCellInput(event) {
  const field = event.target.dataset.field;
  const rowIndex = Number(event.target.dataset.row);
  if (!disposedRows[rowIndex]) {
    return;
  }
  disposedRows[rowIndex][field] = event.target.value;
  invalidDisposedCells.delete(`${rowIndex}||${field}`);
  warnDisposedCells.delete(`${rowIndex}||${field}`);
  validationStateDisposed.delete(rowIndex);
}

function handleDisposedCellChange(event) {
  const field = event.target.dataset.field;
  const rowIndex = Number(event.target.dataset.row);
  if (!disposedRows[rowIndex]) {
    return;
  }
  disposedRows[rowIndex][field] = event.target.value;
  invalidDisposedCells.delete(`${rowIndex}||${field}`);
  warnDisposedCells.delete(`${rowIndex}||${field}`);
  validationStateDisposed.delete(rowIndex);
  renderDisposedAssetsTable();
}

function handleServiceCriteriaCellInput(event) {
  const field = event.target.dataset.field;
  const rowIndex = Number(event.target.dataset.row);
  if (!serviceCriteriaRows[rowIndex]) {
    return;
  }
  serviceCriteriaRows[rowIndex][field] = event.target.value;
}

function handleServiceCriteriaCellChange(event) {
  const field = event.target.dataset.field;
  const rowIndex = Number(event.target.dataset.row);
  if (!serviceCriteriaRows[rowIndex]) {
    return;
  }
  serviceCriteriaRows[rowIndex][field] = event.target.value;
  renderServiceCriteriaTable();
}

function resolveServiceCriteriaRow(rowData, headerMap) {
  const options = {};
  const serviceLookup = lookupTables.serviceCriteria;
  if (!serviceLookup?.objects?.length) {
    return { options };
  }

  const assetIdField = headerMap.get(normalizeHeader("Asset ID"));
  const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
  const compNameField = headerMap.get(normalizeHeader("Component Name"));
  const compTypeField = headerMap.get(normalizeHeader("Component Type"));
  const criteriaField = headerMap.get(normalizeHeader("Service Criteria Type"));

  const subtypeValue = normalizeValue(rowData[subtypeField]);
  const compNameValue = normalizeValue(rowData[compNameField]);
  const compTypeValue = normalizeValue(rowData[compTypeField]);

  const addOptions = (field, values) => {
    if (!field) {
      return;
    }
    const unique = Array.from(new Set(values.filter(Boolean)));
    if (!unique.length) {
      return;
    }
    unique.sort((a, b) => a.localeCompare(b));
    options[field] = unique;
  };

  const allCandidates = serviceLookup.objects;
  addOptions(subtypeField, allCandidates.map((candidate) => candidate.assetsubtype));
  if (assetIdField) {
    addOptions(assetIdField, []);
  }

  let filtered = allCandidates;
  if (subtypeValue) {
    filtered = filtered.filter(
      (candidate) => normalizeValue(candidate.assetsubtype) === subtypeValue
    );
  }

  let compTypeCandidates = filtered;
  if (compNameValue) {
    compTypeCandidates = compTypeCandidates.filter(
      (candidate) => normalizeValue(candidate.componentname) === compNameValue
    );
  }
  addOptions(compTypeField, compTypeCandidates.map((candidate) => candidate.componenttype));

  let compNameCandidates = filtered;
  if (compTypeValue) {
    compNameCandidates = compNameCandidates.filter(
      (candidate) => normalizeValue(candidate.componenttype) === compTypeValue
    );
  }
  addOptions(compNameField, compNameCandidates.map((candidate) => candidate.componentname));

  let criteriaCandidates = filtered;
  if (compTypeValue) {
    criteriaCandidates = criteriaCandidates.filter(
      (candidate) => normalizeValue(candidate.componenttype) === compTypeValue
    );
  }
  if (compNameValue) {
    criteriaCandidates = criteriaCandidates.filter(
      (candidate) => normalizeValue(candidate.componentname) === compNameValue
    );
  }
  addOptions(
    criteriaField,
    criteriaCandidates.map((candidate) => candidate.servicecriteriatype)
  );

  return { options };
}

function applyServiceCriteriaDefault(rowData, headerMap) {
  const serviceLookup = lookupTables.serviceCriteria;
  if (!serviceLookup?.objects?.length) {
    return;
  }
  const assetIdField = headerMap.get(normalizeHeader("Asset ID"));
  const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
  const compNameField = headerMap.get(normalizeHeader("Component Name"));
  const compTypeField = headerMap.get(normalizeHeader("Component Type"));
  const criteriaField = headerMap.get(normalizeHeader("Service Criteria Type"));
  if (!criteriaField || String(rowData[criteriaField] || "").trim()) {
    return;
  }
  const subtypeValue = normalizeValue(rowData[subtypeField]);
  if (!subtypeValue) {
    return;
  }
  let candidates = serviceLookup.objects.filter(
    (candidate) => normalizeValue(candidate.assetsubtype) === subtypeValue
  );
  const compNameValue = normalizeValue(rowData[compNameField]);
  const compTypeValue = normalizeValue(rowData[compTypeField]);
  if (compTypeValue) {
    candidates = candidates.filter(
      (candidate) => normalizeValue(candidate.componenttype) === compTypeValue
    );
  }
  if (compNameValue) {
    candidates = candidates.filter(
      (candidate) => normalizeValue(candidate.componentname) === compNameValue
    );
  }
  const types = Array.from(
    new Set(candidates.map((candidate) => candidate.servicecriteriatype).filter(Boolean))
  );
  if (!types.length) {
    return;
  }
  const main = types.find((type) => normalizeValue(type) === "main condition");
  if (main) {
    rowData[criteriaField] = main;
  } else if (types.length === 1) {
    rowData[criteriaField] = types[0];
  }
}

function isServiceCriteriaRowEmpty(rowData) {
  return serviceCriteriaHeaders.every((header) => {
    const normalized = normalizeHeader(header);
    const value = String(rowData[header] || "").trim();
    if (!value) {
      return true;
    }
    if (normalized === normalizeHeader("Assessed By Resource Name")) {
      return normalizeValue(value) === normalizeValue("CapEx Project");
    }
    if (normalized === normalizeHeader("Score")) {
      return normalizeValue(value) === normalizeValue("0");
    }
    if (normalized === normalizeHeader("Project Code")) {
      return true;
    }
    return false;
  });
}

function syncServiceCriteriaFromSources() {
  const headerMap = buildServiceCriteriaHeaders();
  const newHeaderMap = buildNewAssetHeaders();
  const renewedHeaderMap = buildRenewedHeaders();
  const assetIdNewField =
    newHeaderMap.get(normalizeHeader("Asset ID (if known)")) ||
    newHeaderMap.get(normalizeHeader("Asset ID"));
  const subtypeNewField = newHeaderMap.get(normalizeHeader("Asset SubType"));
  const compNameNewField = newHeaderMap.get(normalizeHeader("Component Name"));
  const compTypeNewField = newHeaderMap.get(normalizeHeader("Component Type"));
  const valuationDateNewField = newHeaderMap.get(normalizeHeader("Valuation Date"));
  const assetIdRenewedField = renewedHeaderMap.get(normalizeHeader("Asset ID"));
  const subtypeRenewedField = renewedHeaderMap.get(normalizeHeader("Asset SubType"));
  const compNameRenewedField = renewedHeaderMap.get(normalizeHeader("Component Name"));
  const compTypeRenewedField = renewedHeaderMap.get(normalizeHeader("Component Type"));
  const valuationDateRenewedField = renewedHeaderMap.get(normalizeHeader("Valuation Date"));
  const conditionRatingRenewedField = renewedHeaderMap.get(normalizeHeader("Condition Rating"));
  const assetIdServiceField = headerMap.get(normalizeHeader("Asset ID"));
  const assessmentDateField = headerMap.get(normalizeHeader("Assessment Date"));
  const scoreField = headerMap.get(normalizeHeader("Score"));
  const assessedByField = headerMap.get(normalizeHeader("Assessed By Resource Name"));

  const sources = new Map();
  const addSource = (assetId, subtype, compName, compType, valuationDate, score, sourceType) => {
    if (!String(subtype || "").trim()) {
      return;
    }
    const key = buildServiceCriteriaKey(assetId, subtype, compName, compType);
    if (!key || sources.has(key)) {
      return;
    }
    sources.set(key, { assetId, subtype, compName, compType, valuationDate, score, sourceType });
  };

  newRows.forEach((rowData) => {
    addSource(
      rowData[assetIdNewField],
      rowData[subtypeNewField],
      rowData[compNameNewField],
      rowData[compTypeNewField],
      rowData[valuationDateNewField],
      null,
      "new"
    );
  });
  renewedRows.forEach((rowData) => {
    addSource(
      rowData[assetIdRenewedField],
      rowData[subtypeRenewedField],
      rowData[compNameRenewedField],
      rowData[compTypeRenewedField],
      rowData[valuationDateRenewedField],
      rowData[conditionRatingRenewedField],
      "renewed"
    );
  });

  const existing = new Map();
  serviceCriteriaRows.forEach((rowData) => {
    const key = rowData.__key || buildServiceCriteriaKey(
      rowData[assetIdServiceField],
      rowData[headerMap.get(normalizeHeader("Asset SubType"))],
      rowData[headerMap.get(normalizeHeader("Component Name"))],
      rowData[headerMap.get(normalizeHeader("Component Type"))]
    );
    if (key) {
      rowData.__key = key;
      existing.set(key, rowData);
    }
  });

  let added = 0;
  sources.forEach((source, key) => {
    const existingRow = existing.get(key);
    if (existingRow) {
      if (assetIdServiceField && !String(existingRow[assetIdServiceField] || "").trim()) {
        existingRow[assetIdServiceField] = source.assetId;
      }
      if (assessmentDateField && !String(existingRow[assessmentDateField] || "").trim()) {
        existingRow[assessmentDateField] = source.valuationDate || "";
      }
      if (scoreField) {
        if (source.sourceType === "new" && !String(existingRow[scoreField] || "").trim()) {
          existingRow[scoreField] = "0";
        } else if (source.sourceType === "renewed" && String(source.score || "").trim()) {
          existingRow[scoreField] = String(source.score).trim();
        }
      }
      if (assessedByField && !String(existingRow[assessedByField] || "").trim()) {
        existingRow[assessedByField] = "CapEx Project";
      }
      const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
      const compNameField = headerMap.get(normalizeHeader("Component Name"));
      const compTypeField = headerMap.get(normalizeHeader("Component Type"));
      if (subtypeField && !String(existingRow[subtypeField] || "").trim()) {
        existingRow[subtypeField] = source.subtype;
      }
      if (compNameField && !String(existingRow[compNameField] || "").trim()) {
        existingRow[compNameField] = source.compName;
      }
      if (compTypeField && !String(existingRow[compTypeField] || "").trim()) {
        existingRow[compTypeField] = source.compType;
      }
      applyServiceCriteriaDefault(existingRow, headerMap);
      return;
    }
    const rowData = buildServiceCriteriaEmptyRow();
    rowData.__key = key;
    if (assetIdServiceField) {
      rowData[assetIdServiceField] = source.assetId || "";
    }
    if (assessmentDateField) {
      rowData[assessmentDateField] = source.valuationDate || "";
    }
    if (scoreField) {
      if (source.sourceType === "new") {
        rowData[scoreField] = "0";
      } else if (source.sourceType === "renewed" && String(source.score || "").trim()) {
        rowData[scoreField] = String(source.score).trim();
      }
    }
    if (assessedByField) {
      rowData[assessedByField] = "CapEx Project";
    }
    const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
    const compNameField = headerMap.get(normalizeHeader("Component Name"));
    const compTypeField = headerMap.get(normalizeHeader("Component Type"));
    if (subtypeField) {
      rowData[subtypeField] = source.subtype || "";
    }
    if (compNameField) {
      rowData[compNameField] = source.compName || "";
    }
    if (compTypeField) {
      rowData[compTypeField] = source.compType || "";
    }
    applyServiceCriteriaDefault(rowData, headerMap);
    serviceCriteriaRows.push(rowData);
    added += 1;
  });

  for (let i = serviceCriteriaRows.length - 1; i >= 0; i -= 1) {
    const rowData = serviceCriteriaRows[i];
    if (isServiceCriteriaRowEmpty(rowData)) {
      serviceCriteriaRows.splice(i, 1);
    }
  }

  renderServiceCriteriaTable();
}

function resolveRenewedRow(rowData, headerMap) {
  const options = {};
  const resolvedFields = new Set();
  const ambiguousFields = new Set();

  const assetIdField = headerMap.get(normalizeHeader("Asset ID"));
  const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
  const typeField = headerMap.get(normalizeHeader("Asset Type"));
  const subClassField = headerMap.get(normalizeHeader("Asset SubClass"));
  const classField = headerMap.get(normalizeHeader("Asset Class"));
  const compNameField = headerMap.get(normalizeHeader("Component Name"));
  const compTypeField = headerMap.get(normalizeHeader("Component Type"));
  const valuationCompField = headerMap.get(normalizeHeader("Valuation Component Name"));

  const addOptions = (field, values) => {
    if (!field) {
      return;
    }
    const unique = Array.from(new Set(values.filter(Boolean)));
    if (!unique.length) {
      return;
    }
    unique.sort((a, b) => a.localeCompare(b));
    options[field] = unique;
    if (!String(rowData[field] || "").trim()) {
      if (unique.length > 1) {
        ambiguousFields.add(field);
      } else if (unique.length === 1) {
        setValueIfEmpty(rowData, field, unique[0], resolvedFields);
      }
    }
  };

  const assetIdValue = normalizeValue(rowData[assetIdField]);
  let valuationCandidates = [];
  if (assetIdValue && ensureValuationIndex()) {
    valuationCandidates = lookupTables.valuationComponent.indexByAssetId.get(assetIdValue) || [];
    if (valuationCandidates.length === 1) {
      const match = valuationCandidates[0];
      setValueIfEmpty(rowData, subtypeField, match.assetsubtype, resolvedFields);
    } else if (valuationCandidates.length > 1) {
      const uniqueSubtypes = Array.from(
        new Set(valuationCandidates.map((candidate) => candidate.assetsubtype).filter(Boolean))
      );
      if (uniqueSubtypes.length === 1) {
        setValueIfEmpty(rowData, subtypeField, uniqueSubtypes[0], resolvedFields);
      } else {
        addOptions(subtypeField, uniqueSubtypes);
      }
    }
  }

  const subtypeValue = normalizeValue(rowData[subtypeField]);
  if (subtypeValue && valuationCandidates.length) {
    let filtered = valuationCandidates.filter(
      (candidate) => normalizeValue(candidate.assetsubtype) === subtypeValue
    );
    const currentCompName = normalizeValue(rowData[compNameField]);
    const currentCompType = normalizeValue(rowData[compTypeField]);
    const currentValComp = normalizeValue(rowData[valuationCompField]);

    if (currentCompName) {
      filtered = filtered.filter((candidate) => normalizeValue(candidate.componentname) === currentCompName);
    }
    if (currentCompType) {
      filtered = filtered.filter((candidate) => normalizeValue(candidate.componenttype) === currentCompType);
    }
    if (currentValComp) {
      filtered = filtered.filter(
        (candidate) => normalizeValue(candidate.valuationcomponentname) === currentValComp
      );
    }

    if (filtered.length === 1) {
      const match = filtered[0];
      setValueIfEmpty(rowData, compNameField, match.componentname, resolvedFields);
      setValueIfEmpty(rowData, compTypeField, match.componenttype, resolvedFields);
      setValueIfEmpty(rowData, valuationCompField, match.valuationcomponentname, resolvedFields);
    } else if (filtered.length > 1) {
      const typeCandidates = filtered;
      addOptions(compTypeField, typeCandidates.map((candidate) => candidate.componenttype));
      const nameCandidates = currentCompType
        ? filtered.filter((candidate) => normalizeValue(candidate.componenttype) === currentCompType)
        : filtered;
      if (currentCompType && !currentCompName) {
        const uniqueNames = Array.from(
          new Set(nameCandidates.map((candidate) => candidate.componentname).filter(Boolean))
        );
        if (uniqueNames.length === 1) {
          setValueIfEmpty(rowData, compNameField, uniqueNames[0], resolvedFields);
        }
      }
      addOptions(compNameField, nameCandidates.map((candidate) => candidate.componentname));
      const valuationCandidatesFiltered = currentCompName || currentCompType
        ? filtered.filter((candidate) => {
            if (currentCompName && normalizeValue(candidate.componentname) !== currentCompName) {
              return false;
            }
            if (currentCompType && normalizeValue(candidate.componenttype) !== currentCompType) {
              return false;
            }
            return true;
          })
        : filtered;
      addOptions(
        valuationCompField,
        valuationCandidatesFiltered.map((candidate) => candidate.valuationcomponentname)
      );
    }
  }

  if (
    !valuationCandidates.length &&
    lookupTables.valuationComponent?.objects?.length &&
    (subtypeValue ||
      String(rowData[compNameField] || "").trim() ||
      String(rowData[compTypeField] || "").trim() ||
      String(rowData[valuationCompField] || "").trim())
  ) {
    const currentSubtype = normalizeValue(rowData[subtypeField]);
    const currentCompName = normalizeValue(rowData[compNameField]);
    const currentCompType = normalizeValue(rowData[compTypeField]);
    const currentValComp = normalizeValue(rowData[valuationCompField]);
    let candidates = lookupTables.valuationComponent.objects;
    if (currentSubtype) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetsubtype) === currentSubtype);
    }
    if (currentCompName) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.componentname) === currentCompName);
    }
    if (currentCompType) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.componenttype) === currentCompType);
    }
    if (currentValComp) {
      candidates = candidates.filter(
        (candidate) => normalizeValue(candidate.valuationcomponentname) === currentValComp
      );
    }

    if (candidates.length === 1) {
      const match = candidates[0];
      setValueIfEmpty(rowData, subtypeField, match.assetsubtype, resolvedFields);
      setValueIfEmpty(rowData, compNameField, match.componentname, resolvedFields);
      setValueIfEmpty(rowData, compTypeField, match.componenttype, resolvedFields);
      setValueIfEmpty(rowData, valuationCompField, match.valuationcomponentname, resolvedFields);
    } else if (candidates.length > 1) {
      if (!currentSubtype) {
        addOptions(subtypeField, candidates.map((candidate) => candidate.assetsubtype));
      }
      addOptions(compTypeField, candidates.map((candidate) => candidate.componenttype));
      const nameCandidates = currentCompType
        ? candidates.filter((candidate) => normalizeValue(candidate.componenttype) === currentCompType)
        : candidates;
      if (currentCompType && !currentCompName) {
        const uniqueNames = Array.from(
          new Set(nameCandidates.map((candidate) => candidate.componentname).filter(Boolean))
        );
        if (uniqueNames.length === 1) {
          setValueIfEmpty(rowData, compNameField, uniqueNames[0], resolvedFields);
        }
      }
      addOptions(compNameField, nameCandidates.map((candidate) => candidate.componentname));
      const valuationCandidatesFiltered = currentCompName || currentCompType
        ? candidates.filter((candidate) => {
            if (currentCompName && normalizeValue(candidate.componentname) !== currentCompName) {
              return false;
            }
            if (currentCompType && normalizeValue(candidate.componenttype) !== currentCompType) {
              return false;
            }
            return true;
          })
        : candidates;
      addOptions(
        valuationCompField,
        valuationCandidatesFiltered.map((candidate) => candidate.valuationcomponentname)
      );
    }
  }

  const assetLookup = lookupTables.assetLookup;
  if (assetLookup?.index) {
    const currentSubtype = normalizeValue(rowData[subtypeField]);
    const currentType = normalizeValue(rowData[typeField]);
    const currentSubClass = normalizeValue(rowData[subClassField]);
    const currentClass = normalizeValue(rowData[classField]);

    let candidates = assetLookup.objects;
    if (currentSubtype) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetsubtype) === currentSubtype);
    }
    if (currentType) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assettype) === currentType);
    }
    if (currentSubClass) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetsubclass) === currentSubClass);
    }
    if (currentClass) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetclass) === currentClass);
    }

    const candidateValues = {
      [classField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assetclass).filter(Boolean))
      ),
      [subClassField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assetsubclass).filter(Boolean))
      ),
      [typeField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assettype).filter(Boolean))
      ),
      [subtypeField]: Array.from(
        new Set(candidates.map((candidate) => candidate.assetsubtype).filter(Boolean))
      )
    };

    [classField, subClassField, typeField, subtypeField].forEach((field) => {
      if (!field) {
        return;
      }
      const values = candidateValues[field] || [];
      addOptions(field, values);
    });

    if (candidates.length === 1) {
      const match = candidates[0];
      setValueIfEmpty(rowData, classField, match.assetclass, resolvedFields);
      setValueIfEmpty(rowData, subClassField, match.assetsubclass, resolvedFields);
      setValueIfEmpty(rowData, typeField, match.assettype, resolvedFields);
      setValueIfEmpty(rowData, subtypeField, match.assetsubtype, resolvedFields);
    }
  }

  const componentLookup = lookupTables.componentLookup;
  if (componentLookup?.index && subtypeValue) {
    const compCandidates = componentLookup.index.get(subtypeValue) || [];
    if (compCandidates.length) {
      const currentType = normalizeValue(rowData[compTypeField]);
      const currentName = normalizeValue(rowData[compNameField]);
      let filtered = compCandidates;
      if (currentType) {
        filtered = filtered.filter((candidate) => normalizeValue(candidate.componenttype) === currentType);
      }
      if (currentName) {
        filtered = filtered.filter((candidate) => normalizeValue(candidate.componentname) === currentName);
      }

      if (filtered.length === 1) {
        const match = filtered[0];
        setValueIfEmpty(rowData, compTypeField, match.componenttype, resolvedFields);
        setValueIfEmpty(rowData, compNameField, match.componentname, resolvedFields);
      } else if (filtered.length > 1) {
        if (!currentType) {
          addOptions(compTypeField, compCandidates.map((candidate) => candidate.componenttype));
        }
        if (currentType && !currentName) {
          addOptions(
            compNameField,
            compCandidates
              .filter((candidate) => normalizeValue(candidate.componenttype) === currentType)
              .map((candidate) => candidate.componentname)
          );
        }
      }
    }
  }

  return { options, resolvedFields, ambiguousFields };
}

function createRenewedCellInput(
  field,
  value,
  options,
  rowIndex,
  colIndex,
  resolvedFields,
  ambiguousFields,
  forceSelect,
  addFillDown = false
) {
  const td = document.createElement("td");
  const isAmbiguous = ambiguousFields.has(field);
  const isResolved = resolvedFields.has(field);
  const isComponentNameField = normalizeHeader(field) === normalizeHeader("Component Name");

  if (isAmbiguous) {
    td.classList.add("cell-ambiguous");
  } else if (isResolved) {
    td.classList.add("cell-resolved");
  }

  const useSelect = !isComponentNameField && (forceSelect || (options && options.length));
  const controlWrap = document.createElement("div");
  controlWrap.className = "cell-control";

  if (useSelect) {
    const select = document.createElement("select");
    select.className = "cell-select";
    select.dataset.field = field;
    select.dataset.row = rowIndex;
    select.dataset.col = colIndex;

    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "";
    select.appendChild(blank);

    const optionValues = Array.isArray(options) ? [...options] : [];
    if (value && !optionValues.includes(value)) {
      optionValues.unshift(value);
    }
    optionValues.forEach((optionValue) => {
      const option = document.createElement("option");
      option.value = optionValue;
      option.textContent = optionValue;
      select.appendChild(option);
    });

    select.value = value || "";
    select.addEventListener("change", handleRenewedCellChange);
    controlWrap.appendChild(select);
  } else {
    const input = document.createElement("input");
    input.type = isDateField(field) ? "date" : "text";
    input.className = "cell-input";
    input.value = value || "";
    input.dataset.field = field;
    input.dataset.row = rowIndex;
    input.dataset.col = colIndex;
    if (isComponentNameField) {
      input.readOnly = true;
    } else {
      input.addEventListener("input", handleRenewedCellInput);
      input.addEventListener("change", handleRenewedCellChange);
    }
    controlWrap.appendChild(input);
  }

  if (addFillDown) {
    const fillBtn = document.createElement("button");
    fillBtn.type = "button";
    fillBtn.className = "field-filldown";
    fillBtn.textContent = "⇩";
    fillBtn.title = `Fill down ${field}`;
    fillBtn.setAttribute("aria-label", `Fill down ${field}`);
    fillBtn.addEventListener("click", () => {
      fillDownRenewedDateField(field);
    });
    controlWrap.appendChild(fillBtn);
  }

  td.appendChild(controlWrap);
  return td;
}

function createDisposedCellInput(
  field,
  value,
  options,
  rowIndex,
  colIndex,
  resolvedFields,
  ambiguousFields,
  forceSelect,
  addFillDown = false
) {
  const td = document.createElement("td");
  const isAmbiguous = ambiguousFields.has(field);
  const isResolved = resolvedFields.has(field);
  const isComponentNameField = normalizeHeader(field) === normalizeHeader("Component Name");

  if (isAmbiguous) {
    td.classList.add("cell-ambiguous");
  } else if (isResolved) {
    td.classList.add("cell-resolved");
  }

  const useSelect = !isComponentNameField && (forceSelect || (options && options.length));
  const controlWrap = document.createElement("div");
  controlWrap.className = "cell-control";

  if (useSelect) {
    const select = document.createElement("select");
    select.className = "cell-select";
    select.dataset.field = field;
    select.dataset.row = rowIndex;
    select.dataset.col = colIndex;

    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "";
    select.appendChild(blank);

    const optionValues = Array.isArray(options) ? [...options] : [];
    if (value && !optionValues.includes(value)) {
      optionValues.unshift(value);
    }
    optionValues.forEach((optionValue) => {
      const option = document.createElement("option");
      option.value = optionValue;
      option.textContent = optionValue;
      select.appendChild(option);
    });

    select.value = value || "";
    select.addEventListener("change", handleDisposedCellChange);
    controlWrap.appendChild(select);
  } else {
    const input = document.createElement("input");
    input.type = isDateField(field) ? "date" : "text";
    input.className = "cell-input";
    input.value = value || "";
    input.dataset.field = field;
    input.dataset.row = rowIndex;
    input.dataset.col = colIndex;
    if (isComponentNameField) {
      input.readOnly = true;
    } else {
      input.addEventListener("input", handleDisposedCellInput);
      input.addEventListener("change", handleDisposedCellChange);
    }
    controlWrap.appendChild(input);
  }

  if (addFillDown) {
    const fillBtn = document.createElement("button");
    fillBtn.type = "button";
    fillBtn.className = "field-filldown";
    fillBtn.textContent = "⇩";
    fillBtn.title = "Fill down Valuation Date";
    fillBtn.setAttribute("aria-label", "Fill down Valuation Date");
    fillBtn.addEventListener("click", () => {
      fillDownDisposedDateField(field);
    });
    controlWrap.appendChild(fillBtn);
  }

  td.appendChild(controlWrap);
  return td;
}

function createServiceCriteriaCellInput(field, value, options, rowIndex, colIndex, forceSelect) {
  const td = document.createElement("td");
  const useSelect = forceSelect || (options && options.length);
  const controlWrap = document.createElement("div");
  controlWrap.className = "cell-control";

  if (useSelect) {
    const select = document.createElement("select");
    select.className = "cell-select";
    select.dataset.field = field;
    select.dataset.row = rowIndex;
    select.dataset.col = colIndex;

    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "";
    select.appendChild(blank);

    const optionValues = Array.isArray(options) ? [...options] : [];
    if (value && !optionValues.includes(value)) {
      optionValues.unshift(value);
    }
    optionValues.forEach((optionValue) => {
      const option = document.createElement("option");
      option.value = optionValue;
      option.textContent = optionValue;
      select.appendChild(option);
    });

    select.value = value || "";
    select.addEventListener("change", handleServiceCriteriaCellChange);
    controlWrap.appendChild(select);
  } else {
    const input = document.createElement("input");
    input.type = "text";
    input.className = "cell-input";
    input.value = value || "";
    input.dataset.field = field;
    input.dataset.row = rowIndex;
    input.dataset.col = colIndex;
    input.addEventListener("input", handleServiceCriteriaCellInput);
    controlWrap.appendChild(input);
  }

  td.appendChild(controlWrap);
  return td;
}

function renderNewAssetsTable() {
  console.log("[debug] renderNewAssetsTable", {
    activeTag: document.activeElement?.tagName,
    activeField: document.activeElement?.dataset?.field,
    activeRow: document.activeElement?.dataset?.row,
    activeCol: document.activeElement?.dataset?.col,
    time: new Date().toISOString()
  });
  const activeElement = document.activeElement;
  let focusSnapshot = null;
  if (activeElement && newAssetsTable.contains(activeElement)) {
    const field = activeElement.dataset.field;
    const row = activeElement.dataset.row;
    const col = activeElement.dataset.col;
    if (field && row != null && col != null) {
      focusSnapshot = {
        field,
        row,
        col,
        tagName: activeElement.tagName.toLowerCase(),
        selectionStart: activeElement.selectionStart,
        selectionEnd: activeElement.selectionEnd
      };
    }
  }

  const headerMap = buildNewAssetHeaders();
  const displayHeaders = [
    headerMap.get(normalizeHeader("Asset Category")),
    ...newHeaders.filter(
      (header) => normalizeHeader(header) !== normalizeHeader("Asset Category")
    )
  ].filter(Boolean);
  newAssetsTable.innerHTML = "";

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const statusHeader = document.createElement("th");
  statusHeader.classList.add("status-col");
  statusHeader.textContent = "Ready";
  headerRow.appendChild(statusHeader);
  const selectHeader = document.createElement("th");
  selectHeader.classList.add("select-col");
  selectHeader.textContent = "Select";
  headerRow.appendChild(selectHeader);
  displayHeaders.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  newAssetsTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  newRows.forEach((rowData, rowIndex) => {
    const { options, resolvedFields, ambiguousFields } = resolveNewAssetRow(rowData, headerMap);
    const tr = document.createElement("tr");
    if (selectedRowIndices.has(rowIndex)) {
      tr.classList.add("row-selected");
    }
    const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
    const subtypeValue = normalizeValue(rowData[subtypeField]);

    const statusCell = document.createElement("td");
    statusCell.classList.add("status-col");
    const statusLight = document.createElement("span");
    const isComplete = isLookupComplete(rowData, headerMap);
    statusLight.className = `status-light ${isComplete ? "status-light--ok" : "status-light--warn"}`;
    statusLight.setAttribute(
      "aria-label",
      isComplete ? "All lookup inputs complete" : "Lookup inputs incomplete"
    );
    statusCell.appendChild(statusLight);
    tr.appendChild(statusCell);

    const selectCell = document.createElement("td");
    selectCell.classList.add("select-col");
    const selectInput = document.createElement("input");
    selectInput.type = "checkbox";
    selectInput.checked = selectedRowIndices.has(rowIndex);
    selectInput.addEventListener("change", () => {
      if (selectInput.checked) {
        selectedRowIndices.add(rowIndex);
      } else {
        selectedRowIndices.delete(rowIndex);
      }
      renderNewAssetsTable();
    });
    selectCell.appendChild(selectInput);
    tr.appendChild(selectCell);

    displayHeaders.forEach((header, colIndex) => {
      const normalized = normalizeHeader(header);
      let fieldOptions = options[header] || null;
      const forceSelect = FORCE_SELECT_FIELDS.has(header);

      if (normalized === normalizeHeader(subtypeField) && !fieldOptions && subtypeOptions.length) {
        fieldOptions = subtypeOptions;
      }

      const addFillDown = rowIndex === 0 && isDateField(header);
      tr.appendChild(
        createCellInput(
          header,
          rowData[header],
          fieldOptions,
          rowIndex,
          colIndex,
          headerMap,
          resolvedFields,
          ambiguousFields,
          forceSelect || normalized === normalizeHeader(subtypeField),
          addFillDown
        )
      );
    });

    tbody.appendChild(tr);
  });
  newAssetsTable.appendChild(tbody);

  if (focusSnapshot) {
    requestAnimationFrame(() => {
      const selector = `${focusSnapshot.tagName}[data-row="${focusSnapshot.row}"][data-col="${focusSnapshot.col}"]`;
      const next = newAssetsTable.querySelector(selector);
      if (!next) {
        return;
      }
      next.focus();
      if (focusSnapshot.tagName === "input" &&
        focusSnapshot.selectionStart != null &&
        focusSnapshot.selectionEnd != null) {
        next.setSelectionRange(focusSnapshot.selectionStart, focusSnapshot.selectionEnd);
      }
    });
  }
}

function renderRenewedAssetsTable() {
  if (!renewedAssetsTable) {
    return;
  }
  const activeElement = document.activeElement;
  let focusSnapshot = null;
  if (activeElement && renewedAssetsTable.contains(activeElement)) {
    const field = activeElement.dataset.field;
    const row = activeElement.dataset.row;
    const col = activeElement.dataset.col;
    if (field && row != null && col != null) {
      focusSnapshot = {
        field,
        row,
        col,
        tagName: activeElement.tagName.toLowerCase(),
        selectionStart: activeElement.selectionStart,
        selectionEnd: activeElement.selectionEnd
      };
    }
  }
  const headerMap = buildRenewedHeaders();
  const displayHeaders = renewedHeaders
    .filter(Boolean)
    .filter((header) => {
      if (!hideRenewedNonCore) {
        return true;
      }
      const normalized = normalizeHeader(header);
      if (RENEWED_HIDE_OVERRIDES.has(normalized)) {
        return false;
      }
      return isCoreField(header);
    });
  renewedAssetsTable.innerHTML = "";

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const statusHeader = document.createElement("th");
  statusHeader.classList.add("status-col");
  statusHeader.textContent = "Ready";
  headerRow.appendChild(statusHeader);
  const selectHeader = document.createElement("th");
  selectHeader.classList.add("select-col");
  selectHeader.textContent = "Select";
  headerRow.appendChild(selectHeader);
  displayHeaders.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  renewedAssetsTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  renewedRows.forEach((rowData, rowIndex) => {
    const tr = document.createElement("tr");
    const { options, resolvedFields, ambiguousFields } = resolveRenewedRow(rowData, headerMap);
    const statusCell = document.createElement("td");
    statusCell.classList.add("status-col");
    const statusLight = document.createElement("span");
    const validationState = validationStateRenewed.get(rowIndex);
    const isComplete = isLookupComplete(rowData, headerMap);
    let statusClass = isComplete ? "status-light--ok" : "status-light--warn";
    if (validationState === "error") {
      statusClass = "status-light--error";
    } else if (validationState === "warn") {
      statusClass = "status-light--warn";
    }
    statusLight.className = `status-light ${statusClass}`;
    statusLight.setAttribute(
      "aria-label",
      isComplete ? "All lookup inputs complete" : "Lookup inputs incomplete"
    );
    statusCell.appendChild(statusLight);
    tr.appendChild(statusCell);

    const selectCell = document.createElement("td");
    selectCell.classList.add("select-col");
    const selectInput = document.createElement("input");
    selectInput.type = "checkbox";
    selectInput.checked = selectedRenewedRowIndices.has(rowIndex);
    selectInput.addEventListener("change", () => {
      if (selectInput.checked) {
        selectedRenewedRowIndices.add(rowIndex);
      } else {
        selectedRenewedRowIndices.delete(rowIndex);
      }
      renderRenewedAssetsTable();
    });
    selectCell.appendChild(selectInput);
    tr.appendChild(selectCell);

    displayHeaders.forEach((header, colIndex) => {
      const normalized = normalizeHeader(header);
      let fieldOptions = options[header] || null;
      const forceSelect = FORCE_SELECT_FIELDS.has(header);
      if (normalized === normalizeHeader("Asset SubType") && !fieldOptions && subtypeOptions.length) {
        fieldOptions = subtypeOptions;
      }
      const addFillDown = rowIndex === 0 && isDateField(header);
      const cell = createRenewedCellInput(
        header,
        rowData[header],
        fieldOptions,
        rowIndex,
        colIndex,
        resolvedFields,
        ambiguousFields,
        forceSelect || normalized === normalizeHeader("Asset SubType"),
        addFillDown
      );
      if (invalidRenewedCells.has(`${rowIndex}||${header}`)) {
        cell.classList.add("cell-invalid");
      } else if (warnRenewedCells.has(`${rowIndex}||${header}`)) {
        cell.classList.add("cell-warning");
      }
      tr.appendChild(cell);
    });
    tbody.appendChild(tr);
  });
  renewedAssetsTable.appendChild(tbody);

  if (focusSnapshot) {
    requestAnimationFrame(() => {
      const selector = `${focusSnapshot.tagName}[data-row="${focusSnapshot.row}"][data-col="${focusSnapshot.col}"]`;
      const next = renewedAssetsTable.querySelector(selector);
      if (!next) {
        return;
      }
      next.focus();
      if (
        focusSnapshot.tagName === "input" &&
        focusSnapshot.selectionStart != null &&
        focusSnapshot.selectionEnd != null
      ) {
        next.setSelectionRange(focusSnapshot.selectionStart, focusSnapshot.selectionEnd);
      }
    });
  }

  if (renewedStatus) {
    renewedStatus.textContent = renewedRows.length
      ? `Rows: ${renewedRows.length}`
      : "No Asset IDs loaded.";
  }
}

function renderDisposedAssetsTable() {
  if (!disposedAssetsTable) {
    return;
  }
  const activeElement = document.activeElement;
  let focusSnapshot = null;
  if (activeElement && disposedAssetsTable.contains(activeElement)) {
    const field = activeElement.dataset.field;
    const row = activeElement.dataset.row;
    const col = activeElement.dataset.col;
    if (field && row != null && col != null) {
      focusSnapshot = {
        field,
        row,
        col,
        tagName: activeElement.tagName.toLowerCase(),
        selectionStart: activeElement.selectionStart,
        selectionEnd: activeElement.selectionEnd
      };
    }
  }
  const headerMap = buildDisposedHeaders();
  const displayHeaders = disposedHeaders
    .filter(Boolean)
    .filter((header) => {
      if (!hideDisposedNonCore) {
        return true;
      }
      const normalized = normalizeHeader(header);
      if (DISPOSED_HIDE_OVERRIDES.has(normalized)) {
        return false;
      }
      return isCoreField(header);
    });
  disposedAssetsTable.innerHTML = "";

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const statusHeader = document.createElement("th");
  statusHeader.classList.add("status-col");
  statusHeader.textContent = "Ready";
  headerRow.appendChild(statusHeader);
  const selectHeader = document.createElement("th");
  selectHeader.classList.add("select-col");
  selectHeader.textContent = "Select";
  headerRow.appendChild(selectHeader);
  displayHeaders.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  disposedAssetsTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  disposedRows.forEach((rowData, rowIndex) => {
    const tr = document.createElement("tr");
    const { options, resolvedFields, ambiguousFields } = resolveRenewedRow(rowData, headerMap);
    const statusCell = document.createElement("td");
    statusCell.classList.add("status-col");
    const statusLight = document.createElement("span");
    const validationState = validationStateDisposed.get(rowIndex);
    const isComplete = isLookupComplete(rowData, headerMap);
    let statusClass = isComplete ? "status-light--ok" : "status-light--warn";
    if (validationState === "error") {
      statusClass = "status-light--error";
    } else if (validationState === "warn") {
      statusClass = "status-light--warn";
    }
    statusLight.className = `status-light ${statusClass}`;
    statusLight.setAttribute(
      "aria-label",
      isComplete ? "All lookup inputs complete" : "Lookup inputs incomplete"
    );
    statusCell.appendChild(statusLight);
    tr.appendChild(statusCell);

    const selectCell = document.createElement("td");
    selectCell.classList.add("select-col");
    const selectInput = document.createElement("input");
    selectInput.type = "checkbox";
    selectInput.checked = selectedDisposedRowIndices.has(rowIndex);
    selectInput.addEventListener("change", () => {
      if (selectInput.checked) {
        selectedDisposedRowIndices.add(rowIndex);
      } else {
        selectedDisposedRowIndices.delete(rowIndex);
      }
      renderDisposedAssetsTable();
    });
    selectCell.appendChild(selectInput);
    tr.appendChild(selectCell);

    displayHeaders.forEach((header, colIndex) => {
      const normalized = normalizeHeader(header);
      let fieldOptions = options[header] || null;
      const forceSelect = FORCE_SELECT_FIELDS.has(header);
      const addFillDown = rowIndex === 0 && isDateField(header);
      if (normalized === normalizeHeader("Asset SubType") && !fieldOptions && subtypeOptions.length) {
        fieldOptions = subtypeOptions;
      }
      const cell = createDisposedCellInput(
        header,
        rowData[header],
        fieldOptions,
        rowIndex,
        colIndex,
        resolvedFields,
        ambiguousFields,
        forceSelect || normalized === normalizeHeader("Asset SubType"),
        addFillDown
      );
      if (invalidDisposedCells.has(`${rowIndex}||${header}`)) {
        cell.classList.add("cell-invalid");
      } else if (warnDisposedCells.has(`${rowIndex}||${header}`)) {
        cell.classList.add("cell-warning");
      }
      tr.appendChild(cell);
    });
    tbody.appendChild(tr);
  });
  disposedAssetsTable.appendChild(tbody);

  if (focusSnapshot) {
    requestAnimationFrame(() => {
      const selector = `${focusSnapshot.tagName}[data-row="${focusSnapshot.row}"][data-col="${focusSnapshot.col}"]`;
      const next = disposedAssetsTable.querySelector(selector);
      if (!next) {
        return;
      }
      next.focus();
      if (
        focusSnapshot.tagName === "input" &&
        focusSnapshot.selectionStart != null &&
        focusSnapshot.selectionEnd != null
      ) {
        next.setSelectionRange(focusSnapshot.selectionStart, focusSnapshot.selectionEnd);
      }
    });
  }

  if (disposedStatus) {
    disposedStatus.textContent = disposedRows.length
      ? `Rows: ${disposedRows.length}`
      : "No Asset IDs loaded.";
  }
}

function renderServiceCriteriaTable() {
  if (!serviceCriteriaTable) {
    return;
  }
  const headerMap = buildServiceCriteriaHeaders();
  const displayHeaders = serviceCriteriaHeaders.filter(Boolean);
  serviceCriteriaTable.innerHTML = "";

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const selectHeader = document.createElement("th");
  selectHeader.classList.add("select-col");
  selectHeader.textContent = "Select";
  headerRow.appendChild(selectHeader);
  displayHeaders.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  serviceCriteriaTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  serviceCriteriaRows.forEach((rowData, rowIndex) => {
    const tr = document.createElement("tr");
    applyServiceCriteriaDefault(rowData, headerMap);
    const { options } = resolveServiceCriteriaRow(rowData, headerMap);

    const selectCell = document.createElement("td");
    selectCell.classList.add("select-col");
    const selectInput = document.createElement("input");
    selectInput.type = "checkbox";
    selectInput.checked = selectedServiceCriteriaRowIndices.has(rowIndex);
    selectInput.addEventListener("change", () => {
      if (selectInput.checked) {
        selectedServiceCriteriaRowIndices.add(rowIndex);
      } else {
        selectedServiceCriteriaRowIndices.delete(rowIndex);
      }
      renderServiceCriteriaTable();
    });
    selectCell.appendChild(selectInput);
    tr.appendChild(selectCell);

    displayHeaders.forEach((header, colIndex) => {
      const normalized = normalizeHeader(header);
      const fieldOptions = options[header] || null;
      const forceSelect = SERVICE_CRITERIA_SELECT_FIELDS.has(normalized);
      tr.appendChild(
        createServiceCriteriaCellInput(
          header,
          rowData[header],
          fieldOptions,
          rowIndex,
          colIndex,
          forceSelect
        )
      );
    });
    tbody.appendChild(tr);
  });
  serviceCriteriaTable.appendChild(tbody);

  if (serviceCriteriaStatus) {
    serviceCriteriaStatus.textContent = serviceCriteriaRows.length
      ? `Rows: ${serviceCriteriaRows.length}`
      : "Ready.";
  }
}

function handleCellChange(event) {
  const field = event.target.dataset.field;
  const rowIndex = Number(event.target.dataset.row);
  const value = event.target.value;
  if (!newRows[rowIndex]) {
    return;
  }
  newRows[rowIndex][field] = value;
  renderNewAssetsTable();
  syncServiceCriteriaFromSources();
}

function addNewRow() {
  if (newRows.length >= MAX_ROWS) {
    newAssetStatus.textContent = `Row limit reached (${MAX_ROWS}).`;
    return;
  }
  const rowData = buildEmptyRow();
  newRows.push(rowData);
  newAssetStatus.textContent = `Rows: ${newRows.length}`;
  renderNewAssetsTable();
}

function addServiceCriteriaRow() {
  if (serviceCriteriaRows.length >= MAX_ROWS) {
    if (serviceCriteriaStatus) {
      serviceCriteriaStatus.textContent = `Row limit reached (${MAX_ROWS}).`;
    }
    return;
  }
  const rowData = buildServiceCriteriaEmptyRow();
  serviceCriteriaRows.push(rowData);
  renderServiceCriteriaTable();
}

function insertServiceCriteriaRowAt(index, rowData) {
  if (serviceCriteriaRows.length >= MAX_ROWS) {
    if (serviceCriteriaStatus) {
      serviceCriteriaStatus.textContent = `Row limit reached (${MAX_ROWS}).`;
    }
    return;
  }
  const targetIndex = Math.max(0, Math.min(index, serviceCriteriaRows.length));
  if (selectedServiceCriteriaRowIndices.size) {
    const shifted = new Set();
    selectedServiceCriteriaRowIndices.forEach((value) => {
      shifted.add(value >= targetIndex ? value + 1 : value);
    });
    selectedServiceCriteriaRowIndices.clear();
    shifted.forEach((value) => selectedServiceCriteriaRowIndices.add(value));
  }
  serviceCriteriaRows.splice(targetIndex, 0, rowData);
  selectedServiceCriteriaRowIndices.add(targetIndex);
  renderServiceCriteriaTable();
}

function cloneServiceCriteriaRow(rowData) {
  const clone = cloneRowData(rowData);
  delete clone.__key;
  return clone;
}

function duplicateLatestServiceCriteriaRows() {
  if (!serviceCriteriaRows.length) {
    if (serviceCriteriaStatus) {
      serviceCriteriaStatus.textContent = "No rows to duplicate.";
    }
    return;
  }
  const count = Number(serviceCriteriaDuplicateCountSelect?.value) || 1;
  const baseRow = serviceCriteriaRows[serviceCriteriaRows.length - 1];
  for (let i = 0; i < count; i += 1) {
    if (serviceCriteriaRows.length >= MAX_ROWS) {
      break;
    }
    serviceCriteriaRows.push(cloneServiceCriteriaRow(baseRow));
  }
  selectedServiceCriteriaRowIndices.clear();
  selectedServiceCriteriaRowIndices.add(serviceCriteriaRows.length - 1);
  renderServiceCriteriaTable();
}

function insertServiceCriteriaRow() {
  if (!selectedServiceCriteriaRowIndices.size) {
    insertServiceCriteriaRowAt(serviceCriteriaRows.length, buildServiceCriteriaEmptyRow());
    return;
  }
  const indices = Array.from(selectedServiceCriteriaRowIndices).sort((a, b) => a - b);
  let offset = 0;
  indices.forEach((index) => {
    if (serviceCriteriaRows.length >= MAX_ROWS) {
      return;
    }
    insertServiceCriteriaRowAt(index + offset, buildServiceCriteriaEmptyRow());
    offset += 1;
  });
}

function insertDuplicateServiceCriteriaRow() {
  if (!selectedServiceCriteriaRowIndices.size) {
    if (serviceCriteriaStatus) {
      serviceCriteriaStatus.textContent = "Select a row to insert above.";
    }
    return;
  }
  const indices = Array.from(selectedServiceCriteriaRowIndices).sort((a, b) => a - b);
  const baseRows = indices.map((index) => cloneServiceCriteriaRow(serviceCriteriaRows[index]));
  let offset = 0;
  indices.forEach((index, idx) => {
    if (serviceCriteriaRows.length >= MAX_ROWS) {
      return;
    }
    insertServiceCriteriaRowAt(index + offset, baseRows[idx]);
    offset += 1;
  });
}

function removeSelectedServiceCriteriaRows() {
  if (!serviceCriteriaRows.length || !selectedServiceCriteriaRowIndices.size) {
    if (serviceCriteriaStatus) {
      serviceCriteriaStatus.textContent = "Select rows to remove.";
    }
    return;
  }
  const indices = Array.from(selectedServiceCriteriaRowIndices).sort((a, b) => b - a);
  indices.forEach((index) => {
    if (serviceCriteriaRows[index]) {
      serviceCriteriaRows.splice(index, 1);
    }
  });
  selectedServiceCriteriaRowIndices.clear();
  if (serviceCriteriaRows.length) {
    renderServiceCriteriaTable();
  }
}

function updateAllRows() {
  if (newRows.length) {
    renderNewAssetsTable();
  }
  if (renewedRows.length) {
    renderRenewedAssetsTable();
  }
  if (disposedRows.length) {
    renderDisposedAssetsTable();
  }
  syncServiceCriteriaFromSources();
}

function getAllowedValuesForValidation(rowData, headerMap) {
  const allowed = new Map();
  const addAllowed = (field, values) => {
    if (!field) {
      return;
    }
    const unique = Array.from(new Set(values.filter(Boolean)));
    if (!unique.length) {
      return;
    }
    allowed.set(field, new Set(unique));
  };

  const assetCategoryField = headerMap.get(normalizeHeader("Asset Category"));
  const assetClassField = headerMap.get(normalizeHeader("Asset Class"));
  const assetSubClassField = headerMap.get(normalizeHeader("Asset SubClass"));
  const assetTypeField = headerMap.get(normalizeHeader("Asset Type"));
  const assetSubtypeField = headerMap.get(normalizeHeader("Asset SubType"));
  const compNameField = headerMap.get(normalizeHeader("Component Name"));
  const compTypeField = headerMap.get(normalizeHeader("Component Type"));
  const assetIdField = headerMap.get(normalizeHeader("Asset ID"));

  const assetLookup = lookupTables.assetLookup;
  if (assetLookup?.objects?.length) {
    const currentSubtype = normalizeValue(rowData[assetSubtypeField]);
    const currentType = normalizeValue(rowData[assetTypeField]);
    const currentSubClass = normalizeValue(rowData[assetSubClassField]);
    const currentClass = normalizeValue(rowData[assetClassField]);
    const currentCategory = normalizeValue(rowData[assetCategoryField]);
    let candidates = assetLookup.objects;
    if (currentSubtype) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetsubtype) === currentSubtype);
    }
    if (currentType) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assettype) === currentType);
    }
    if (currentSubClass) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetsubclass) === currentSubClass);
    }
    if (currentClass) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetclass) === currentClass);
    }
    if (currentCategory) {
      candidates = candidates.filter((candidate) => normalizeValue(candidate.assetcategory) === currentCategory);
    }
    addAllowed(assetCategoryField, candidates.map((candidate) => candidate.assetcategory));
    addAllowed(assetClassField, candidates.map((candidate) => candidate.assetclass));
    addAllowed(assetSubClassField, candidates.map((candidate) => candidate.assetsubclass));
    addAllowed(assetTypeField, candidates.map((candidate) => candidate.assettype));
    addAllowed(assetSubtypeField, candidates.map((candidate) => candidate.assetsubtype));
  }

  const subtypeValue = normalizeValue(rowData[assetSubtypeField]);
  const componentLookup = lookupTables.componentLookup;
  if (componentLookup?.index && subtypeValue) {
    const compCandidates = componentLookup.index.get(subtypeValue) || [];
    addAllowed(compTypeField, compCandidates.map((candidate) => candidate.componenttype));
    const currentType = normalizeValue(rowData[compTypeField]);
    const filteredByType = currentType
      ? compCandidates.filter((candidate) => normalizeValue(candidate.componenttype) === currentType)
      : compCandidates;
    addAllowed(compNameField, filteredByType.map((candidate) => candidate.componentname));
  }

  const valuationLookup = lookupTables.valuationComponent;
  if (valuationLookup?.indexByAssetId && assetIdField) {
    const assetIdValue = normalizeValue(rowData[assetIdField]);
    if (assetIdValue) {
      const matches = valuationLookup.indexByAssetId.get(assetIdValue) || [];
      let filtered = matches;
      if (subtypeValue) {
        filtered = filtered.filter((candidate) => normalizeValue(candidate.assetsubtype) === subtypeValue);
      }
      addAllowed(compTypeField, filtered.map((candidate) => candidate.componenttype));
      const currentType = normalizeValue(rowData[compTypeField]);
      const filteredByType = currentType
        ? filtered.filter((candidate) => normalizeValue(candidate.componenttype) === currentType)
        : filtered;
      addAllowed(compNameField, filteredByType.map((candidate) => candidate.componentname));
    }
  }

  return allowed;
}

function validateLookupRows(rows, headerMap, invalidSet, warnSet, stateMap) {
  invalidSet.clear();
  warnSet.clear();
  stateMap.clear();
  ensureValuationIndex();
  let invalidCount = 0;
  let warnCount = 0;
  let validatedFields = 0;
  rows.forEach((rowData, rowIndex) => {
    const allowed = getAllowedValuesForValidation(rowData, headerMap);
    const compNameField = headerMap.get(normalizeHeader("Component Name"));
    const compTypeField = headerMap.get(normalizeHeader("Component Type"));
    const assetIdField = headerMap.get(normalizeHeader("Asset ID"));
    const assetIdValue = assetIdField ? normalizeValue(rowData[assetIdField]) : "";
    if (assetIdValue && lookupTables.valuationComponent?.indexByAssetId) {
      const matches = lookupTables.valuationComponent.indexByAssetId.get(assetIdValue) || [];
      if (matches.length) {
        const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
        const subtypeValue = normalizeValue(rowData[subtypeField]);
        let filtered = matches;
        if (subtypeValue) {
          filtered = filtered.filter((candidate) => normalizeValue(candidate.assetsubtype) === subtypeValue);
        }
        if (compTypeField && !allowed.has(compTypeField)) {
          allowed.set(
            compTypeField,
            new Set(filtered.map((candidate) => candidate.componenttype).filter(Boolean))
          );
        }
        if (compNameField && !allowed.has(compNameField)) {
          const currentType = normalizeValue(rowData[compTypeField]);
          const filteredByType = currentType
            ? filtered.filter((candidate) => normalizeValue(candidate.componenttype) === currentType)
            : filtered;
          allowed.set(
            compNameField,
            new Set(filteredByType.map((candidate) => candidate.componentname).filter(Boolean))
          );
        }
      }
    }
    let rowHasError = false;
    let rowHasWarn = false;
    const deferred = [];
    allowed.forEach((allowedSet, field) => {
      if (field === compNameField) {
        deferred.push([field, allowedSet]);
        return;
      }
      const value = String(rowData[field] || "").trim();
      if (!value) {
        return;
      }
      validatedFields += 1;
      const normalizedValue = normalizeValue(value);
      const normalizedAllowed = new Set(Array.from(allowedSet).map((item) => normalizeValue(item)));
      if (!normalizedAllowed.has(normalizedValue)) {
        invalidCount += 1;
        rowHasError = true;
        invalidSet.add(`${rowIndex}||${field}`);
      }
    });
    deferred.forEach(([field, allowedSet]) => {
      const value = String(rowData[field] || "").trim();
      if (!value) {
        return;
      }
      validatedFields += 1;
      if (compTypeField && invalidSet.has(`${rowIndex}||${compTypeField}`)) {
        return;
      }
      const normalizedValue = normalizeValue(value);
      const normalizedAllowed = new Set(Array.from(allowedSet).map((item) => normalizeValue(item)));
      if (!normalizedAllowed.has(normalizedValue)) {
        invalidCount += 1;
        rowHasError = true;
        invalidSet.add(`${rowIndex}||${field}`);
      }
    });
    if (rowHasError) {
      stateMap.set(rowIndex, "error");
    } else if (rowHasWarn) {
      stateMap.set(rowIndex, "warn");
    } else {
      stateMap.set(rowIndex, "ok");
    }
  });
  return { invalidCount, warnCount, validatedFields };
}

function fillDownDisposedValuationDate() {
  const headerMap = buildDisposedHeaders();
  const field = headerMap.get(normalizeHeader("Valuation Date"));
  if (!field || !disposedRows.length) {
    return;
  }
  const value = disposedRows[0]?.[field];
  if (!String(value || "").trim()) {
    if (disposedStatus) {
      disposedStatus.textContent = "Enter a Valuation Date in the first row before fill-down.";
    }
    return;
  }
  disposedRows.forEach((rowData, rowIndex) => {
    rowData[field] = value;
    invalidDisposedCells.delete(`${rowIndex}||${field}`);
    warnDisposedCells.delete(`${rowIndex}||${field}`);
    validationStateDisposed.delete(rowIndex);
  });
  renderDisposedAssetsTable();
}

function fillDownNewDateField(field) {
  if (!newRows.length) {
    return;
  }
  const value = newRows[0]?.[field];
  if (!String(value || "").trim()) {
    if (newAssetStatus) {
      newAssetStatus.textContent = `Enter ${field} in the first row before fill-down.`;
    }
    return;
  }
  newRows.forEach((rowData) => {
    rowData[field] = value;
  });
  renderNewAssetsTable();
}

function fillDownRenewedDateField(field) {
  if (!renewedRows.length) {
    return;
  }
  const value = renewedRows[0]?.[field];
  if (!String(value || "").trim()) {
    if (renewedStatus) {
      renewedStatus.textContent = `Enter ${field} in the first row before fill-down.`;
    }
    return;
  }
  renewedRows.forEach((rowData, rowIndex) => {
    rowData[field] = value;
    invalidRenewedCells.delete(`${rowIndex}||${field}`);
    warnRenewedCells.delete(`${rowIndex}||${field}`);
    validationStateRenewed.delete(rowIndex);
  });
  renderRenewedAssetsTable();
}

function fillDownDisposedDateField(field) {
  if (!disposedRows.length) {
    return;
  }
  const value = disposedRows[0]?.[field];
  if (!String(value || "").trim()) {
    if (disposedStatus) {
      disposedStatus.textContent = `Enter ${field} in the first row before fill-down.`;
    }
    return;
  }
  disposedRows.forEach((rowData, rowIndex) => {
    rowData[field] = value;
    invalidDisposedCells.delete(`${rowIndex}||${field}`);
    warnDisposedCells.delete(`${rowIndex}||${field}`);
    validationStateDisposed.delete(rowIndex);
  });
  renderDisposedAssetsTable();
}

function buildEmptyRow() {
  const rowData = {};
  newHeaders.forEach((header) => {
    rowData[header] = NEW_ROW_DEFAULTS[header] || "";
  });
  if (projectCodeValue) {
    const projectHeader = newHeaders.find(
      (header) => normalizeHeader(header) === normalizeHeader("Project Code")
    );
    if (projectHeader) {
      rowData[projectHeader] = projectCodeValue;
    }
  }
  return rowData;
}

function cloneRowData(rowData) {
  return { ...rowData };
}

function insertRowAt(index, rowData) {
  if (newRows.length >= MAX_ROWS) {
    newAssetStatus.textContent = `Row limit reached (${MAX_ROWS}).`;
    return;
  }
  const targetIndex = Math.max(0, Math.min(index, newRows.length));
  if (selectedRowIndices.size) {
    const shifted = new Set();
    selectedRowIndices.forEach((value) => {
      shifted.add(value >= targetIndex ? value + 1 : value);
    });
    selectedRowIndices.clear();
    shifted.forEach((value) => selectedRowIndices.add(value));
  }
  newRows.splice(targetIndex, 0, rowData);
  selectedRowIndices.add(targetIndex);
  newAssetStatus.textContent = `Rows: ${newRows.length}`;
  renderNewAssetsTable();
}

function duplicateLatestRows() {
  if (!newRows.length) {
    newAssetStatus.textContent = "No rows to duplicate.";
    return;
  }
  const count = Number(duplicateCountSelect?.value) || 1;
  const baseRow = newRows[newRows.length - 1];
  for (let i = 0; i < count; i += 1) {
    if (newRows.length >= MAX_ROWS) {
      break;
    }
    newRows.push(cloneRowData(baseRow));
  }
  selectedRowIndices.clear();
  selectedRowIndices.add(newRows.length - 1);
  newAssetStatus.textContent = `Rows: ${newRows.length}`;
  renderNewAssetsTable();
}

function insertEmptyRow() {
  if (!selectedRowIndices.size) {
    insertRowAt(newRows.length, buildEmptyRow());
    return;
  }
  const indices = Array.from(selectedRowIndices).sort((a, b) => a - b);
  let offset = 0;
  indices.forEach((index) => {
    if (newRows.length >= MAX_ROWS) {
      return;
    }
    insertRowAt(index + offset, buildEmptyRow());
    offset += 1;
  });
}

function insertDuplicateRow() {
  if (!selectedRowIndices.size) {
    newAssetStatus.textContent = "Select a row to insert above.";
    return;
  }
  const indices = Array.from(selectedRowIndices).sort((a, b) => a - b);
  const baseRows = indices.map((index) => cloneRowData(newRows[index]));
  let offset = 0;
  indices.forEach((index, idx) => {
    if (newRows.length >= MAX_ROWS) {
      return;
    }
    insertRowAt(index + offset, baseRows[idx]);
    offset += 1;
  });
}

function removeSelectedRows() {
  if (!selectedRowIndices.size) {
    newAssetStatus.textContent = "Select rows to remove.";
    return;
  }
  const indices = Array.from(selectedRowIndices).sort((a, b) => b - a);
  indices.forEach((index) => {
    if (newRows[index]) {
      newRows.splice(index, 1);
    }
  });
  selectedRowIndices.clear();
  if (!newRows.length) {
    addNewRow();
  } else {
    newAssetStatus.textContent = `Rows: ${newRows.length}`;
    renderNewAssetsTable();
  }
}

function resetNewAssets() {
  newRows = [];
  selectedRowIndices.clear();
  addNewRow();
  newAssetResult.textContent = "";
}

function resetRenewedAssets() {
  renewedRows = [];
  selectedRenewedRowIndices.clear();
  invalidRenewedCells.clear();
  renderRenewedAssetsTable();
}

function exportNewCsv() {
  const exportHeaders = newHeaders.filter(
    (header) => normalizeHeader(header) !== normalizeHeader("Asset Category")
  );
  const rows = newRows.map((row) => exportHeaders.map((header) => row[header] || ""));
  const csv = toCSV(exportHeaders, rows);
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = "new_assets_interactive.csv";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function exportRenewedCsv() {
  const exportHeaders = renewedHeaders.filter(Boolean);
  const rows = renewedRows.map((row) => exportHeaders.map((header) => row[header] || ""));
  const csv = toCSV(exportHeaders, rows);
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = "renewed_assets_interactive.csv";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function exportDisposedCsv() {
  const exportHeaders = disposedHeaders.filter(Boolean);
  const rows = disposedRows.map((row) => exportHeaders.map((header) => row[header] || ""));
  const csv = toCSV(exportHeaders, rows);
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = "disposed_assets_interactive.csv";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function exportServiceCriteriaCsv() {
  const exportHeaders = serviceCriteriaHeaders.filter(Boolean);
  const rows = serviceCriteriaRows.map((row) => exportHeaders.map((header) => row[header] || ""));
  const csv = toCSV(exportHeaders, rows);
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = "service_criteria_interactive.csv";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function buildSessionData() {
  return {
    version: 1,
    exportedAt: new Date().toISOString(),
    projectCode: projectCodeValue,
    headers: {
      new: newHeaders,
      renewed: renewedHeaders,
      disposed: disposedHeaders,
      serviceCriteria: serviceCriteriaHeaders
    },
    rows: {
      new: newRows,
      renewed: renewedRows,
      disposed: disposedRows,
      serviceCriteria: serviceCriteriaRows
    }
  };
}

function downloadSession() {
  if (projectCodeInput) {
    projectCodeValue = projectCodeInput.value.trim();
  }
  const data = buildSessionData();
  const json = JSON.stringify(data, null, 2);
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  const prefix = sanitizeFilePart(projectCodeValue);
  link.download = prefix ? `${prefix}_par_session.json` : "par_session.json";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function applySessionData(data) {
  if (!data || typeof data !== "object") {
    throw new Error("Invalid session file.");
  }

  newHeaders = Array.isArray(data.headers?.new) ? data.headers.new : [...DEFAULT_NEW_HEADERS];
  renewedHeaders = Array.isArray(data.headers?.renewed)
    ? data.headers.renewed
    : [...DEFAULT_RENEWED_HEADERS];
  disposedHeaders = Array.isArray(data.headers?.disposed)
    ? data.headers.disposed
    : [...DEFAULT_DISPOSED_HEADERS];
  serviceCriteriaHeaders = Array.isArray(data.headers?.serviceCriteria)
    ? data.headers.serviceCriteria
    : [...DEFAULT_SERVICE_HEADERS];

  newRows = Array.isArray(data.rows?.new) ? data.rows.new : [];
  renewedRows = Array.isArray(data.rows?.renewed) ? data.rows.renewed : [];
  disposedRows = Array.isArray(data.rows?.disposed) ? data.rows.disposed : [];
  serviceCriteriaRows = Array.isArray(data.rows?.serviceCriteria) ? data.rows.serviceCriteria : [];
  projectCodeValue = String(data.projectCode || "").trim();
  if (projectCodeInput) {
    projectCodeInput.value = projectCodeValue;
  }

  selectedRowIndices.clear();
  selectedRenewedRowIndices.clear();
  selectedDisposedRowIndices.clear();
  selectedServiceCriteriaRowIndices.clear();
  invalidRenewedCells.clear();
  invalidDisposedCells.clear();
  warnRenewedCells.clear();
  warnDisposedCells.clear();
  validationStateRenewed.clear();
  validationStateDisposed.clear();
  subtypeOptions = lookupTables.assetLookup?.index
    ? Array.from(lookupTables.assetLookup.index.keys())
        .map(
          (value) =>
            lookupTables.assetLookup.objects.find(
              (obj) => normalizeValue(obj.assetsubtype) === value
            )?.assetsubtype || value
        )
        .sort((a, b) => a.localeCompare(b))
    : [];

  renderNewAssetsTable();
  renderRenewedAssetsTable();
  renderDisposedAssetsTable();
  syncServiceCriteriaFromSources();
  applyProjectCodeToRows();
  if (!newRows.length) {
    addNewRow();
  }
}

async function handleSessionFile(event) {
  const file = event.target.files[0];
  if (!file) {
    return;
  }
  try {
    const text = await readFileAsText(file);
    const data = JSON.parse(text);
    const loadedCount = Object.values(lookupTables).filter(Boolean).length;
    if (!loadedCount) {
      window.alert(
        "Lookup tables are not loaded. Load lookup CSVs before loading a session."
      );
      return;
    }
    applySessionData(data);
  } catch (error) {
    if (lookupStatus) {
      updateLookupStatus("Failed to load session file.", true);
    }
    console.error(error);
  }
}

async function loadNewHeaders() {
  try {
    const text = await fetchTextWithFallback(HEADER_FILES.new);
    const parsed = parseCSV(text.trim());
    if (parsed.length && parsed[0].length) {
      newHeaders = parsed[0];
    }
  } catch (error) {
    newHeaders = [...DEFAULT_NEW_HEADERS];
  }
}

async function loadRenewedHeaders() {
  try {
    const text = await fetchTextWithFallback(HEADER_FILES.renewed);
    const parsed = parseCSV(text.trim());
    if (parsed.length && parsed[0].length) {
      renewedHeaders = parsed[0];
    }
  } catch (error) {
    renewedHeaders = [...DEFAULT_RENEWED_HEADERS];
  }
}

async function loadDisposedHeaders() {
  try {
    const text = await fetchTextWithFallback(HEADER_FILES.disposed);
    const parsed = parseCSV(text.trim());
    if (parsed.length && parsed[0].length) {
      disposedHeaders = parsed[0];
    }
  } catch (error) {
    disposedHeaders = [...DEFAULT_DISPOSED_HEADERS];
  }
}

async function loadServiceCriteriaHeaders() {
  try {
    const text = await fetchTextWithFallback(HEADER_FILES.serviceCriteria);
    const parsed = parseCSV(text.trim());
    if (parsed.length && parsed[0].length) {
      serviceCriteriaHeaders = parsed[0];
    }
  } catch (error) {
    serviceCriteriaHeaders = [...DEFAULT_SERVICE_HEADERS];
  }
}

function parseAssetIdList(text) {
  const trimmed = String(text || "").trim();
  if (!trimmed) {
    return [];
  }
  const parsed = parseCSV(trimmed);
  if (!parsed.length) {
    return [];
  }
  const headerMap = buildHeaderMap(parsed[0]);
  const idIndex = headerMap.get(normalizeHeader("Asset ID"));
  let values = [];
  if (idIndex != null) {
    values = parsed.slice(1).map((row) => row[idIndex]);
  } else {
    values = parsed.flat();
  }
  values = values.flatMap((value) => {
    const textValue = String(value || "");
    if (!textValue) {
      return [];
    }
    const split = textValue.split(/[\s,;]+/g).filter(Boolean);
    return split.length ? split : [textValue];
  });
  const seen = new Set();
  const ids = [];
  values.forEach((value) => {
    const cleaned = String(value || "").trim();
    if (!cleaned) {
      return;
    }
    const key = cleaned.toLowerCase();
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    ids.push(cleaned);
  });
  return ids;
}

function buildRenewedEmptyRow() {
  const rowData = {};
  renewedHeaders.forEach((header) => {
    rowData[header] = "";
  });
  if (projectCodeValue) {
    const projectHeader = renewedHeaders.find(
      (header) => normalizeHeader(header) === normalizeHeader("Project Code")
    );
    if (projectHeader) {
      rowData[projectHeader] = projectCodeValue;
    }
  }
  return rowData;
}

function buildDisposedEmptyRow() {
  const rowData = {};
  disposedHeaders.forEach((header) => {
    const normalized = normalizeHeader(header);
    if (normalized === normalizeHeader("Valuation Record Type")) {
      rowData[header] = "Full Disposal";
    } else {
      rowData[header] = "";
    }
  });
  if (projectCodeValue) {
    const projectHeader = disposedHeaders.find(
      (header) => normalizeHeader(header) === normalizeHeader("Project Code")
    );
    if (projectHeader) {
      rowData[projectHeader] = projectCodeValue;
    }
  }
  return rowData;
}

function buildServiceCriteriaEmptyRow() {
  const rowData = {};
  serviceCriteriaHeaders.forEach((header) => {
    if (normalizeHeader(header) === normalizeHeader("Assessed By Resource Name")) {
      rowData[header] = "CapEx Project";
    } else {
      rowData[header] = "";
    }
  });
  if (projectCodeValue) {
    const projectHeader = serviceCriteriaHeaders.find(
      (header) => normalizeHeader(header) === normalizeHeader("Project Code")
    );
    if (projectHeader) {
      rowData[projectHeader] = projectCodeValue;
    }
  }
  return rowData;
}

function applyProjectCodeToRow(rowData, headerMap, value) {
  const field = headerMap.get(normalizeHeader("Project Code"));
  if (!field) {
    return;
  }
  rowData[field] = value;
}

function applyProjectCodeToRows() {
  const value = String(projectCodeValue || "").trim();
  if (!value) {
    return;
  }
  const newHeaderMap = buildNewAssetHeaders();
  const renewedHeaderMap = buildRenewedHeaders();
  const disposedHeaderMap = buildDisposedHeaders();
  const serviceHeaderMap = buildServiceCriteriaHeaders();

  newRows.forEach((rowData) => applyProjectCodeToRow(rowData, newHeaderMap, value));
  renewedRows.forEach((rowData) => applyProjectCodeToRow(rowData, renewedHeaderMap, value));
  disposedRows.forEach((rowData) => applyProjectCodeToRow(rowData, disposedHeaderMap, value));
  serviceCriteriaRows.forEach((rowData) => applyProjectCodeToRow(rowData, serviceHeaderMap, value));
}

function fillRenewedFromLookup(rowData, headerMap, match) {
  const resolvedFields = new Set();
  const assetIdField = headerMap.get(normalizeHeader("Asset ID"));
  const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
  const typeField = headerMap.get(normalizeHeader("Asset Type"));
  const subClassField = headerMap.get(normalizeHeader("Asset SubClass"));
  const classField = headerMap.get(normalizeHeader("Asset Class"));
  const compNameField = headerMap.get(normalizeHeader("Component Name"));
  const compTypeField = headerMap.get(normalizeHeader("Component Type"));
  const valuationCompField = headerMap.get(normalizeHeader("Valuation Component Name"));

  if (match) {
    setValueIfEmpty(rowData, assetIdField, match.asset_id || match.assetid, resolvedFields);
    setValueIfEmpty(rowData, subtypeField, match.assetsubtype, resolvedFields);
    setValueIfEmpty(rowData, compNameField, match.componentname, resolvedFields);
    setValueIfEmpty(rowData, compTypeField, match.componenttype, resolvedFields);
    setValueIfEmpty(rowData, valuationCompField, match.valuationcomponentname, resolvedFields);
  }

  const subtypeValue = normalizeValue(rowData[subtypeField]);
  if (subtypeValue && lookupTables.assetLookup?.index) {
    const candidates = lookupTables.assetLookup.index.get(subtypeValue) || [];
    if (candidates.length) {
      const assetMatch = candidates[0];
      setValueIfEmpty(rowData, classField, assetMatch.assetclass, resolvedFields);
      setValueIfEmpty(rowData, subClassField, assetMatch.assetsubclass, resolvedFields);
      setValueIfEmpty(rowData, typeField, assetMatch.assettype, resolvedFields);
      setValueIfEmpty(rowData, subtypeField, assetMatch.assetsubtype, resolvedFields);
    }
  }
}

function populateRenewedFromIds(ids) {
  const headerMap = buildRenewedHeaders();
  const assetIdField = headerMap.get(normalizeHeader("Asset ID"));
  const valuationLookup = lookupTables.valuationComponent;
  let lookupReady = Boolean(valuationLookup?.objects?.length);
  if (lookupReady) {
    lookupReady = ensureValuationIndex();
  }
  renewedRows = [];
  selectedRenewedRowIndices.clear();
  let matchedCount = 0;
  ids.forEach((id) => {
    if (lookupReady) {
      const matches = valuationLookup.indexByAssetId.get(normalizeValue(id)) || [];
      if (matches.length) {
        matchedCount += 1;
        const subtypeValues = Array.from(
          new Set(matches.map((match) => match.assetsubtype).filter(Boolean))
        );
        const compTypes = Array.from(
          new Set(matches.map((match) => match.componenttype).filter(Boolean))
        );
        const subtypeValue = subtypeValues.length === 1 ? subtypeValues[0] : "";
        if (compTypes.length) {
          compTypes.forEach((compType) => {
            const rowData = buildRenewedEmptyRow();
            if (assetIdField) {
              rowData[assetIdField] = id;
            }
            if (subtypeValue) {
              const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
              if (subtypeField) {
                rowData[subtypeField] = subtypeValue;
              }
            }
            const compTypeField = headerMap.get(normalizeHeader("Component Type"));
            if (compTypeField) {
              rowData[compTypeField] = compType;
            }
            const compNameField = headerMap.get(normalizeHeader("Component Name"));
            if (compNameField) {
              const nameMatch = matches.find(
                (match) => normalizeValue(match.componenttype) === normalizeValue(compType)
              );
              if (nameMatch?.componentname) {
                rowData[compNameField] = nameMatch.componentname;
              }
            }
            renewedRows.push(rowData);
          });
          return;
        }
      }
    }
    const rowData = buildRenewedEmptyRow();
    if (assetIdField) {
      rowData[assetIdField] = id;
    }
    renewedRows.push(rowData);
  });
  renderRenewedAssetsTable();
  syncServiceCriteriaFromSources();
  if (renewedStatus && ids.length) {
    if (!lookupReady) {
      if (valuationLookup?.objects?.length) {
        renewedStatus.textContent = "Valuation Components loaded, but Asset ID column not detected.";
      } else {
        renewedStatus.textContent = "Valuation Component lookup not loaded. Load lookups first.";
      }
    } else if (!matchedCount) {
      renewedStatus.textContent = "No valuation matches found for supplied Asset IDs.";
    }
  }
}

function populateDisposedFromIds(ids) {
  const headerMap = buildDisposedHeaders();
  const assetIdField = headerMap.get(normalizeHeader("Asset ID"));
  const valuationLookup = lookupTables.valuationComponent;
  let lookupReady = Boolean(valuationLookup?.objects?.length);
  if (lookupReady) {
    lookupReady = ensureValuationIndex();
  }
  disposedRows = [];
  selectedDisposedRowIndices.clear();
  invalidDisposedCells.clear();
  warnDisposedCells.clear();
  validationStateDisposed.clear();
  let matchedCount = 0;
  ids.forEach((id) => {
    if (lookupReady) {
      const matches = valuationLookup.indexByAssetId.get(normalizeValue(id)) || [];
      if (matches.length) {
        matchedCount += 1;
        const subtypeValues = Array.from(
          new Set(matches.map((match) => match.assetsubtype).filter(Boolean))
        );
        const compTypes = Array.from(
          new Set(matches.map((match) => match.componenttype).filter(Boolean))
        );
        const subtypeValue = subtypeValues.length === 1 ? subtypeValues[0] : "";
        if (compTypes.length) {
          compTypes.forEach((compType) => {
            const rowData = buildDisposedEmptyRow();
            if (assetIdField) {
              rowData[assetIdField] = id;
            }
            if (subtypeValue) {
              const subtypeField = headerMap.get(normalizeHeader("Asset SubType"));
              if (subtypeField) {
                rowData[subtypeField] = subtypeValue;
              }
            }
            const compTypeField = headerMap.get(normalizeHeader("Component Type"));
            if (compTypeField) {
              rowData[compTypeField] = compType;
            }
            const compNameField = headerMap.get(normalizeHeader("Component Name"));
            if (compNameField) {
              const nameMatch = matches.find(
                (match) => normalizeValue(match.componenttype) === normalizeValue(compType)
              );
              if (nameMatch?.componentname) {
                rowData[compNameField] = nameMatch.componentname;
              }
            }
            disposedRows.push(rowData);
          });
          return;
        }
      }
    }
    const rowData = buildDisposedEmptyRow();
    if (assetIdField) {
      rowData[assetIdField] = id;
    }
    disposedRows.push(rowData);
  });
  renderDisposedAssetsTable();
  if (disposedStatus && ids.length) {
    if (!lookupReady) {
      if (valuationLookup?.objects?.length) {
        disposedStatus.textContent = "Valuation Components loaded, but Asset ID column not detected.";
      } else {
        disposedStatus.textContent = "Valuation Component lookup not loaded. Load lookups first.";
      }
    } else if (!matchedCount) {
      disposedStatus.textContent = "No valuation matches found for supplied Asset IDs.";
    }
  }
}

function setActiveTab(tabId) {
  if (!tabId) {
    return;
  }
  tabButtons.forEach((btn) => {
    const isActive = btn.dataset.tab === tabId;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", isActive ? "true" : "false");
  });
  tabPanels.forEach((panel) => {
    const isActive = panel.dataset.tab === tabId;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
    panel.setAttribute("aria-hidden", isActive ? "false" : "true");
  });
}

if (tabButtons.length) {
  tabButtons.forEach((btn) => {
    btn.setAttribute("role", "tab");
    btn.setAttribute("aria-selected", btn.classList.contains("is-active") ? "true" : "false");
    btn.addEventListener("click", () => {
      setActiveTab(btn.dataset.tab);
    });
  });
  const initialTab =
    tabButtons.find((btn) => btn.classList.contains("is-active"))?.dataset.tab || "lookup";
  setActiveTab(initialTab);
}

async function init() {
  try {
    document.querySelectorAll("details.card").forEach((detail) => {
      detail.open = true;
    });
    if (duplicateCountSelect) {
      duplicateCountSelect.innerHTML = "";
      for (let i = 1; i <= 99; i += 1) {
        const option = document.createElement("option");
        option.value = String(i);
        option.textContent = String(i);
        duplicateCountSelect.appendChild(option);
      }
      duplicateCountSelect.value = "1";
    }
    if (serviceCriteriaDuplicateCountSelect) {
      serviceCriteriaDuplicateCountSelect.innerHTML = "";
      for (let i = 1; i <= 99; i += 1) {
        const option = document.createElement("option");
        option.value = String(i);
        option.textContent = String(i);
        serviceCriteriaDuplicateCountSelect.appendChild(option);
      }
      serviceCriteriaDuplicateCountSelect.value = "1";
    }
    await loadNewHeaders();
    await loadRenewedHeaders();
    await loadDisposedHeaders();
    await loadServiceCriteriaHeaders();
    if (!newRows.length) {
      addNewRow();
    } else {
      renderNewAssetsTable();
    }
    renderRenewedAssetsTable();
    renderDisposedAssetsTable();
    syncServiceCriteriaFromSources();
    newAssetStatus.textContent = `Rows: ${newRows.length}`;
  } catch (error) {
    newAssetResult.textContent = "Table builder failed to initialize. Please refresh.";
    console.error(error);
  }
}

addNewRowBtn.addEventListener("click", addNewRow);
exportNewBtn.addEventListener("click", exportNewCsv);
resetNewBtn.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  resetNewAssets();
});
duplicateNewRowBtn.addEventListener("click", duplicateLatestRows);
insertNewRowBtn.addEventListener("click", insertEmptyRow);
insertDuplicateRowBtn.addEventListener("click", insertDuplicateRow);
removeSelectedRowsBtn.addEventListener("click", removeSelectedRows);
if (addServiceCriteriaRowBtn) {
  addServiceCriteriaRowBtn.addEventListener("click", addServiceCriteriaRow);
}
if (duplicateServiceCriteriaRowBtn) {
  duplicateServiceCriteriaRowBtn.addEventListener("click", duplicateLatestServiceCriteriaRows);
}
if (insertServiceCriteriaRowBtn) {
  insertServiceCriteriaRowBtn.addEventListener("click", insertServiceCriteriaRow);
}
if (insertDuplicateServiceCriteriaRowBtn) {
  insertDuplicateServiceCriteriaRowBtn.addEventListener("click", insertDuplicateServiceCriteriaRow);
}
if (removeServiceCriteriaRowsBtn) {
  removeServiceCriteriaRowsBtn.addEventListener("click", removeSelectedServiceCriteriaRows);
}
if (validateServiceCriteriaBtn) {
  validateServiceCriteriaBtn.addEventListener("click", () => {
    syncServiceCriteriaFromSources();
  });
}
if (exportServiceCriteriaBtn) {
  exportServiceCriteriaBtn.addEventListener("click", exportServiceCriteriaCsv);
}
if (resetServiceCriteriaBtn) {
  resetServiceCriteriaBtn.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    serviceCriteriaRows = [];
    selectedServiceCriteriaRowIndices.clear();
    renderServiceCriteriaTable();
  });
}
if (saveSessionBtn) {
  saveSessionBtn.addEventListener("click", downloadSession);
}
if (loadSessionBtn && sessionFileInput) {
  loadSessionBtn.addEventListener("click", () => {
    const loadedCount = Object.values(lookupTables).filter(Boolean).length;
    if (!loadedCount) {
      window.alert(
        "Lookup tables are not loaded. Load lookup CSVs before loading a session."
      );
      return;
    }
    sessionFileInput.value = "";
    sessionFileInput.click();
  });
  sessionFileInput.addEventListener("change", handleSessionFile);
}
if (applyProjectCodeBtn && projectCodeInput) {
  applyProjectCodeBtn.addEventListener("click", () => {
    projectCodeValue = projectCodeInput.value.trim().toUpperCase();
    projectCodeInput.value = projectCodeValue;
    applyProjectCodeToRows();
    renderNewAssetsTable();
    renderRenewedAssetsTable();
    renderDisposedAssetsTable();
    renderServiceCriteriaTable();
  });
}
if (projectCodeInput) {
  projectCodeInput.addEventListener("input", () => {
    const upper = projectCodeInput.value.toUpperCase();
    if (projectCodeInput.value !== upper) {
      const cursor = projectCodeInput.selectionStart;
      projectCodeInput.value = upper;
      if (cursor != null) {
        projectCodeInput.setSelectionRange(cursor, cursor);
      }
    }
  });
}
if (resetRenewedBtn) {
  resetRenewedBtn.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    resetRenewedAssets();
  });
}
  if (resetDisposedBtn) {
    resetDisposedBtn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      disposedRows = [];
      selectedDisposedRowIndices.clear();
      invalidDisposedCells.clear();
      renderDisposedAssetsTable();
    });
  }
if (renewedLoadIdsBtn && renewedIdFileInput) {
  renewedLoadIdsBtn.addEventListener("click", () => {
    renewedIdFileInput.value = "";
    renewedIdFileInput.click();
  });
  renewedIdFileInput.addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (!file) {
      return;
    }
    const text = await readFileAsText(file);
    const ids = parseAssetIdList(text);
    if (!ids.length && renewedStatus) {
      renewedStatus.textContent = "No Asset IDs found in uploaded file.";
    }
    populateRenewedFromIds(ids);
  });
}
if (renewedApplyIdsBtn && renewedIdText) {
  renewedApplyIdsBtn.addEventListener("click", () => {
    const ids = parseAssetIdList(renewedIdText.value);
    if (!ids.length && renewedStatus) {
      renewedStatus.textContent = "No Asset IDs found in pasted text.";
    }
    populateRenewedFromIds(ids);
  });
}
if (renewedTogglePasteBtn && renewedIdText) {
  renewedTogglePasteBtn.addEventListener("click", () => {
    const isHidden = renewedIdText.classList.toggle("is-hidden");
    renewedTogglePasteBtn.textContent = isHidden ? "Show Paste Box" : "Hide Paste Box";
  });
}
if (exportRenewedBtn) {
  exportRenewedBtn.addEventListener("click", exportRenewedCsv);
}
if (renewedValidateBtn) {
  renewedValidateBtn.addEventListener("click", () => {
    if (!renewedRows.length) {
      if (renewedStatus) {
        renewedStatus.textContent = "No rows to validate.";
      }
      return;
    }
    const headerMap = buildRenewedHeaders();
    const { invalidCount, warnCount, validatedFields } = validateLookupRows(
      renewedRows,
      headerMap,
      invalidRenewedCells,
      warnRenewedCells,
      validationStateRenewed
    );
    renderRenewedAssetsTable();
    if (renewedStatus) {
      if (!validatedFields) {
        renewedStatus.textContent = "No lookup-based fields available for validation.";
      } else if (!invalidCount && !warnCount) {
        renewedStatus.textContent = "Validation complete. No issues found.";
      } else if (invalidCount && warnCount) {
        renewedStatus.textContent = `Validation complete. ${invalidCount} errors and ${warnCount} Component Name warning(s).`;
      } else if (invalidCount) {
        renewedStatus.textContent = `Validation complete. ${invalidCount} errors need review.`;
      } else {
        renewedStatus.textContent = `Validation complete. ${warnCount} Component Name warning(s).`;
      }
    }
  });
}
if (renewedToggleFieldsBtn) {
  renewedToggleFieldsBtn.addEventListener("click", () => {
    hideRenewedNonCore = !hideRenewedNonCore;
    renewedToggleFieldsBtn.textContent = hideRenewedNonCore
      ? "Show All Fields"
      : "Hide Non-Asset Fields";
    renderRenewedAssetsTable();
  });
}

if (disposedLoadIdsBtn && disposedIdFileInput) {
  disposedLoadIdsBtn.addEventListener("click", () => {
    disposedIdFileInput.value = "";
    disposedIdFileInput.click();
  });
  disposedIdFileInput.addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (!file) {
      return;
    }
    const text = await readFileAsText(file);
    const ids = parseAssetIdList(text);
    if (!ids.length && disposedStatus) {
      disposedStatus.textContent = "No Asset IDs found in uploaded file.";
    }
    populateDisposedFromIds(ids);
  });
}
if (disposedApplyIdsBtn && disposedIdText) {
  disposedApplyIdsBtn.addEventListener("click", () => {
    const ids = parseAssetIdList(disposedIdText.value);
    if (!ids.length && disposedStatus) {
      disposedStatus.textContent = "No Asset IDs found in pasted text.";
    }
    populateDisposedFromIds(ids);
  });
}
if (disposedTogglePasteBtn && disposedIdText) {
  disposedTogglePasteBtn.addEventListener("click", () => {
    const isHidden = disposedIdText.classList.toggle("is-hidden");
    disposedTogglePasteBtn.textContent = isHidden ? "Show Paste Box" : "Hide Paste Box";
  });
}
if (exportDisposedBtn) {
  exportDisposedBtn.addEventListener("click", exportDisposedCsv);
}
if (disposedValidateBtn) {
  disposedValidateBtn.addEventListener("click", () => {
    if (!disposedRows.length) {
      if (disposedStatus) {
        disposedStatus.textContent = "No rows to validate.";
      }
      return;
    }
    const headerMap = buildDisposedHeaders();
    const { invalidCount, warnCount, validatedFields } = validateLookupRows(
      disposedRows,
      headerMap,
      invalidDisposedCells,
      warnDisposedCells,
      validationStateDisposed
    );
    renderDisposedAssetsTable();
    if (disposedStatus) {
      if (!validatedFields) {
        disposedStatus.textContent = "No lookup-based fields available for validation.";
      } else if (!invalidCount && !warnCount) {
        disposedStatus.textContent = "Validation complete. No issues found.";
      } else if (invalidCount && warnCount) {
        disposedStatus.textContent = `Validation complete. ${invalidCount} errors and ${warnCount} Component Name warning(s).`;
      } else if (invalidCount) {
        disposedStatus.textContent = `Validation complete. ${invalidCount} errors need review.`;
      } else {
        disposedStatus.textContent = `Validation complete. ${warnCount} Component Name warning(s).`;
      }
    }
  });
}
if (disposedToggleFieldsBtn) {
  disposedToggleFieldsBtn.addEventListener("click", () => {
    hideDisposedNonCore = !hideDisposedNonCore;
    disposedToggleFieldsBtn.textContent = hideDisposedNonCore
      ? "Show All Fields"
      : "Hide Non-Asset Fields";
    renderDisposedAssetsTable();
  });
}

init();

if (toggleWidthBtn) {
  toggleWidthBtn.addEventListener("click", () => {
    const container = document.querySelector(".container");
    if (!container) {
      return;
    }
    const isFull = container.classList.toggle("full-width");
    toggleWidthBtn.textContent = isFull ? "Constrain width" : "Full width";
  });
}

if (resetAllBtn) {
  resetAllBtn.addEventListener("click", () => {
    const confirmed = window.confirm("Reset all tables and clear loaded lookup data?");
    if (!confirmed) {
      return;
    }
    newRows = [];
    renewedRows = [];
    disposedRows = [];
    serviceCriteriaRows = [];
    selectedRowIndices.clear();
    selectedRenewedRowIndices.clear();
    selectedDisposedRowIndices.clear();
    selectedServiceCriteriaRowIndices.clear();
    invalidRenewedCells.clear();
    invalidDisposedCells.clear();
    subtypeOptions = [];
    Object.keys(lookupTables).forEach((key) => {
      lookupTables[key] = null;
    });
    document.querySelectorAll(".lookup-item").forEach((section) => {
      const statusEl = section.querySelector(".lookupStatus");
      const textArea = section.querySelector(".lookupText");
      if (statusEl) {
        statusEl.textContent = "⚠ Not loaded.";
        statusEl.style.color = "";
      }
      if (textArea) {
        textArea.value = "";
        textArea.readOnly = false;
      }
    });
    if (lookupStatus) {
      lookupStatus.textContent = "No lookups loaded.";
      lookupStatus.style.color = "";
    }
    if (renewedIdText) {
      renewedIdText.value = "";
    }
    if (renewedStatus) {
      renewedStatus.textContent = "No Asset IDs loaded.";
    }
    if (disposedIdText) {
      disposedIdText.value = "";
    }
    if (disposedStatus) {
      disposedStatus.textContent = "No Asset IDs loaded.";
    }
    if (serviceCriteriaStatus) {
      serviceCriteriaStatus.textContent = "Ready.";
    }
    newAssetStatus.textContent = "Ready.";
    addNewRow();
    renderRenewedAssetsTable();
    renderDisposedAssetsTable();
    renderServiceCriteriaTable();
  });
}
