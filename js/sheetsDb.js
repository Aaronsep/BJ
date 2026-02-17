const DEFAULT_TTL_MS = 10 * 60 * 1000;
const DEFAULT_STORAGE_KEY = "sheet-db-v1";

function normalizeHeader(value) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function normalizeAddressKey(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function cellToText(cell) {
  if (!cell) return "";
  if (typeof cell.f === "string" && cell.f.trim()) return cell.f.trim();
  if (cell.v === null || cell.v === undefined) return "";
  return String(cell.v).trim();
}

function extractSpreadsheetId(spreadsheetUrl) {
  const value = String(spreadsheetUrl ?? "").trim();
  if (!value) throw new Error("The sheet link is missing");
  const match = value.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  if (/^[a-zA-Z0-9-_]+$/.test(value)) return value;
  throw new Error("The sheet link is not valid");
}

function withQueryParam(urlString, key, val) {
  const url = new URL(urlString);
  if (!url.searchParams.has(key)) {
    url.searchParams.set(key, val);
  }
  return url.toString();
}

function buildEndpointCandidates(spreadsheetUrl) {
  const source = String(spreadsheetUrl ?? "").trim();
  const id = extractSpreadsheetId(source);
  const candidates = [];

  if (source.includes("/gviz/tq")) {
    candidates.push({
      type: "gviz",
      url: withQueryParam(source, "tqx", "out:json"),
      label: "source-gviz",
    });
  }

  if (source.includes("/export") && source.includes("format=csv")) {
    candidates.push({
      type: "csv",
      url: source,
      label: "source-csv",
    });
  }

  candidates.push({
    type: "gviz",
    url: `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:json`,
    label: "id-gviz",
  });
  candidates.push({
    type: "csv",
    url: `https://docs.google.com/spreadsheets/d/${id}/export?format=csv`,
    label: "id-csv",
  });

  return candidates;
}

function parseGvizEnvelope(rawText) {
  const start = rawText.indexOf("{");
  const end = rawText.lastIndexOf("}");
  if (start < 0 || end <= start) throw new Error("Unexpected response format");
  return JSON.parse(rawText.slice(start, end + 1));
}

function parseCsvLine(line) {
  const out = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      const next = line[i + 1];
      if (inQuotes && next === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === "," && !inQuotes) {
      out.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  out.push(current);
  return out.map((v) => v.trim());
}

function parseCsv(rawText) {
  const lines = String(rawText ?? "")
    .replace(/\r/g, "")
    .split("\n")
    .filter((line) => line.length > 0);

  if (!lines.length) return { headers: [], rows: [] };
  const headers = parseCsvLine(lines[0]);
  const rows = lines.slice(1).map(parseCsvLine);
  return { headers, rows };
}

function pickColumnIndexes(headers) {
  const normalized = headers.map(normalizeHeader);
  const addressIdx = normalized.findIndex((h) => h.includes("propertyaddress"));
  const timeIdx = normalized.findIndex((h) => h.includes("scheduledtime"));
  return { addressIdx, timeIdx };
}

function rowsFromGvizTable(envelope) {
  if (envelope.status === "error") {
    throw new Error("The sheet data is currently unavailable");
  }

  const table = envelope.table || {};
  const rows = Array.isArray(table.rows) ? table.rows : [];
  const colLabels = Array.isArray(table.cols) ? table.cols.map((col) => String(col.label || "").trim()) : [];

  let headers = colLabels.slice();
  let startDataRow = 0;

  if (headers.every((h) => !h) && rows.length > 0) {
    const firstCells = Array.isArray(rows[0].c) ? rows[0].c : [];
    headers = firstCells.map(cellToText);
    startDataRow = 1;
  }

  const dataRows = [];
  for (let i = startDataRow; i < rows.length; i++) {
    const cells = Array.isArray(rows[i].c) ? rows[i].c : [];
    dataRows.push(cells.map(cellToText));
  }

  return { headers, rows: dataRows };
}

function toRecords(headers, rows) {
  const { addressIdx, timeIdx } = pickColumnIndexes(headers);
  if (addressIdx < 0 || timeIdx < 0) {
    throw new Error('Required columns were not found: "Property Address" and "Scheduled Time"');
  }

  const dedup = new Map();
  for (const row of rows) {
    const address = String(row[addressIdx] ?? "").trim();
    const scheduledTime = String(row[timeIdx] ?? "").trim();
    if (!address || !scheduledTime) continue;

    dedup.set(normalizeAddressKey(address), { address, scheduledTime });
  }

  return Array.from(dedup.values()).sort((a, b) => a.address.localeCompare(b.address));
}

async function tryEndpoint(candidate) {
  const response = await fetch(candidate.url, { cache: "no-store" });
  if (!response.ok) {
    throw new Error("Could not reach the sheet");
  }

  const rawText = await response.text();
  if (candidate.type === "gviz") {
    const envelope = parseGvizEnvelope(rawText);
    const data = rowsFromGvizTable(envelope);
    return toRecords(data.headers, data.rows);
  }

  const data = parseCsv(rawText);
  return toRecords(data.headers, data.rows);
}

async function fetchSheetRows(spreadsheetUrl) {
  const candidates = buildEndpointCandidates(spreadsheetUrl);
  const errors = [];

  for (const candidate of candidates) {
    try {
      return await tryEndpoint(candidate);
    } catch (error) {
      const reason = error && error.message ? error.message : String(error);
      errors.push(`${candidate.label}: ${reason}`);
    }
  }
  if (errors.length) {
    console.warn("[sheetsDb] Sync failed:", errors.join(" || "));
  }
  throw new Error('Could not sync addresses. Make sure sharing is set to "Anyone with the link (Viewer)".');
}

function createToken() {
  const random = Math.random().toString(36).slice(2, 10);
  return `${Date.now().toString(36)}-${random}`;
}

function safeReadLocalStorage(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    return parsed;
  } catch {
    return null;
  }
}

function safeWriteLocalStorage(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Keep app functional even when storage quota is full.
  }
}

export function createSheetsDb({
  spreadsheetUrl,
  ttlMs = DEFAULT_TTL_MS,
  storageKey = DEFAULT_STORAGE_KEY,
  onStateChange = () => {},
} = {}) {
  const safeTtl = Math.max(60 * 1000, Number(ttlMs) || DEFAULT_TTL_MS);
  const state = {
    records: [],
    token: "",
    refreshedAt: 0,
    expiresAt: 0,
    loading: false,
    lastError: "",
  };

  let refreshPromise = null;
  let refreshTimer = null;

  function emit() {
    onStateChange({
      records: state.records.slice(),
      token: state.token,
      refreshedAt: state.refreshedAt,
      expiresAt: state.expiresAt,
      loading: state.loading,
      lastError: state.lastError,
    });
  }

  function scheduleAutoRefresh() {
    if (refreshTimer) clearTimeout(refreshTimer);
    if (!state.expiresAt) return;

    const msLeft = Math.max(500, state.expiresAt - Date.now());
    refreshTimer = setTimeout(() => {
      void refresh();
    }, msLeft);
  }

  function setFreshData(records) {
    state.records = records;
    state.token = createToken();
    state.refreshedAt = Date.now();
    state.expiresAt = state.refreshedAt + safeTtl;
    state.lastError = "";

    safeWriteLocalStorage(storageKey, {
      records: state.records,
      token: state.token,
      refreshedAt: state.refreshedAt,
      expiresAt: state.expiresAt,
    });
  }

  function loadCache() {
    const cached = safeReadLocalStorage(storageKey);
    if (!cached) return;

    state.records = Array.isArray(cached.records) ? cached.records : [];
    state.token = String(cached.token || "");
    state.refreshedAt = Number(cached.refreshedAt) || 0;
    state.expiresAt = Number(cached.expiresAt) || 0;
  }

  async function refresh() {
    if (refreshPromise) return refreshPromise;

    state.loading = true;
    emit();

    refreshPromise = (async () => {
      try {
        const rows = await fetchSheetRows(spreadsheetUrl);
        setFreshData(rows);
      } catch (error) {
        state.lastError = error && error.message ? error.message : String(error);
        state.expiresAt = Date.now() + 60 * 1000;
      } finally {
        state.loading = false;
        emit();
        scheduleAutoRefresh();
        refreshPromise = null;
      }
    })();

    return refreshPromise;
  }

  function findByAddress(address) {
    const target = normalizeAddressKey(address);
    if (!target) return null;
    return state.records.find((row) => normalizeAddressKey(row.address) === target) || null;
  }

  async function init() {
    loadCache();
    emit();

    const shouldRefresh = !state.records.length || Date.now() >= state.expiresAt;
    if (shouldRefresh) {
      await refresh();
    } else {
      scheduleAutoRefresh();
    }
  }

  function destroy() {
    if (refreshTimer) clearTimeout(refreshTimer);
    refreshTimer = null;
  }

  return {
    init,
    refreshNow: refresh,
    destroy,
    findByAddress,
    getRecords: () => state.records.slice(),
    getState: () => ({
      records: state.records.slice(),
      token: state.token,
      refreshedAt: state.refreshedAt,
      expiresAt: state.expiresAt,
      loading: state.loading,
      lastError: state.lastError,
    }),
  };
}
