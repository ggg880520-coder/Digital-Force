const XLSX = require("xlsx");
const crypto = require("crypto");

/* =========================
   KONFIGURASI
   Disarankan memakai environment variable untuk production:
   SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, SPREADSHEET_ID
========================= */
const SUPABASE_URL = process.env.SUPABASE_URL || "https://wqgkwbtsfrmpgwxhaybz.supabase.co";
const SUPABASE_SERVICE_ROLE_KEY =
  process.env.SUPABASE_SERVICE_ROLE_KEY ||
  process.env.SUPABASE_ANON_KEY ||
  "";
const SPREADSHEET_ID = process.env.SPREADSHEET_ID || "13NWbHrHsyNcDaA1xe2OZPkyoNDuPHO6gLG2Pud3e6hA";

const SHEET_SA = process.env.SHEET_SA || "Rasio SA";
const SHEET_TS = process.env.SHEET_TS || "TimeSheet";

const TABLE_SA = process.env.TABLE_SA || "sa_raw";
const TABLE_TS = process.env.TABLE_TS || "timesheet_raw";

const CHUNK_SIZE = Number(process.env.CHUNK_SIZE || 500);

/* =========================
   HELPERS
========================= */
function asText(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s === "" ? null : s;
}

function asNumber(v) {
  if (v === null || v === undefined || v === "") return null;
  const s = String(v)
    .trim()
    .replace(/[^\d,.-]/g, "")
    .replace(/\.(?=\d{3}(\D|$))/g, "")
    .replace(/,/g, ".");
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function asInteger(v) {
  const n = asNumber(v);
  return Number.isFinite(n) ? Math.round(n) : null;
}

function normalizeHeaderName(header) {
  const raw = String(header ?? "")
    .normalize("NFKC")
    .replace(/\u00A0/g, " ")
    .trim()
    .toLowerCase()
    .replace(/[^\w\s/]/g, "")
    .replace(/\s+/g, "_")
    .replace(/\//g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");

  const aliasMap = {
    case_number: "case_number",
    sa_number: "sa_number",
    actual_dispatched_date: "actual_dispatched_date",
    actual_travel_date: "actual_travel_date",
    actual_inprogress_date: "actual_inprogress_date",
    actual_completed_date: "actual_completed_date",
    timesheet_submitted: "timesheet_submitted",
    time_sheet_submitted: "timesheet_submitted",
    timesheet_approved: "timesheet_approved",
    time_sheet_approved: "timesheet_approved",
    case_description: "case_description",
    service_area: "service_area",
    mechanic_name: "mechanic_name",
    scheduled_start: "scheduled_start",
    sa_accuracy: "sa_accuracy",
    month: "month",
    nrp: "nrp",
    available_hour: "available_hours",
    available_hours: "available_hours",
    qty_sa: "qty_sa",
    ach: "ach",
    tdt: "tdt",
    efh: "efh",
    ja: "ja",
    je: "je",
    ut_partner: "ut_partner",
    date: "date",
    service_resource: "service_resource"
  };

  return aliasMap[raw] || raw;
}

function normalizeRowKeys(row) {
  const out = {};
  for (const [key, value] of Object.entries(row)) {
    out[normalizeHeaderName(key)] = value;
  }
  return out;
}

function hasAnyValue(row) {
  return Object.values(row).some(
    (v) => v !== null && v !== undefined && String(v).trim() !== ""
  );
}

function buildRowHash(obj) {
  return crypto
    .createHash("sha256")
    .update(JSON.stringify(obj))
    .digest("hex");
}

function buildSupabaseUrl(tableName, params = {}) {
  const url = new URL(`/rest/v1/${encodeURIComponent(tableName)}`, SUPABASE_URL);
  for (const [key, value] of Object.entries(params)) {
    url.searchParams.set(key, value);
  }
  return url.toString();
}

async function supabaseRequest(tableName, method, params = {}, body = null, extraHeaders = {}) {
  if (!SUPABASE_SERVICE_ROLE_KEY) {
    throw new Error("SUPABASE_SERVICE_ROLE_KEY belum diisi. Set environment variable sebelum menjalankan npm run sync.");
  }

  const headers = {
    apikey: SUPABASE_SERVICE_ROLE_KEY,
    Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    ...extraHeaders
  };

  if (body !== null) headers["Content-Type"] = "application/json";

  const res = await fetch(buildSupabaseUrl(tableName, params), {
    method,
    headers,
    body: body !== null ? JSON.stringify(body) : undefined
  });

  const text = await res.text();

  if (!res.ok) {
    throw new Error(`Supabase ${method} ${tableName} gagal: ${res.status} ${text}`);
  }

  return text ? JSON.parse(text) : null;
}

async function clearTableBySource(tableName, sourceSheet) {
  await supabaseRequest(
    tableName,
    "DELETE",
    { source_sheet: `eq.${sourceSheet}` },
    null,
    { Prefer: "return=minimal" }
  );
}

async function insertChunks(tableName, rows) {
  if (!rows.length) {
    console.log(`[${tableName}] tidak ada row untuk diinsert`);
    return;
  }

  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    await supabaseRequest(
      tableName,
      "POST",
      {},
      chunk,
      { Prefer: "return=minimal" }
    );
    console.log(`[${tableName}] inserted ${Math.min(i + CHUNK_SIZE, rows.length)}/${rows.length}`);
  }
}

async function downloadWorkbookFromGoogleSheet() {
  const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx`;
  const res = await fetch(url);

  if (!res.ok) {
    throw new Error(`Gagal download spreadsheet: ${res.status} ${res.statusText}`);
  }

  const arr = await res.arrayBuffer();
  return XLSX.read(Buffer.from(arr), {
    type: "buffer",
    cellDates: false
  });
}

function readSheetObjects(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet tidak ditemukan: ${sheetName}`);

  return XLSX.utils.sheet_to_json(ws, {
    defval: null,
    raw: false
  })
    .filter(hasAnyValue)
    .map(normalizeRowKeys);
}

/* =========================
   MAPPING SESUAI KOLOM DASHBOARD
========================= */
function mapRasioSaRows(rows) {
  return rows
    .map((r) => {
      const row = {
        case_number: asText(r.case_number),
        sa_number: asText(r.sa_number),
        actual_dispatched_date: asText(r.actual_dispatched_date),
        actual_travel_date: asText(r.actual_travel_date),
        actual_inprogress_date: asText(r.actual_inprogress_date),
        actual_completed_date: asText(r.actual_completed_date),
        timesheet_submitted: asText(r.timesheet_submitted),
        timesheet_approved: asText(r.timesheet_approved),
        case_description: asText(r.case_description),
        service_area: asText(r.service_area),
        mechanic_name: asText(r.mechanic_name),
        scheduled_start: asText(r.scheduled_start),
        sa_accuracy: asInteger(r.sa_accuracy),
        source_sheet: SHEET_SA
      };
      row.source_row_hash = buildRowHash(row);
      return row;
    })
    .filter((r) => r.case_number || r.sa_number || r.mechanic_name);
}

function mapTimesheetRows(rows) {
  return rows
    .map((r) => {
      const row = {
        month: asText(r.month),
        nrp: asText(r.nrp),
        mechanic_name: asText(r.mechanic_name),
        available_hours: asNumber(r.available_hours),
        qty_sa: asNumber(r.qty_sa),
        ach: asNumber(r.ach),
        tdt: asNumber(r.tdt),
        efh: asNumber(r.efh),
        ja: asNumber(r.ja),
        je: asNumber(r.je),
        ut_partner: asText(r.ut_partner),
        date: asText(r.date),
        service_resource: asText(r.service_resource),
        source_sheet: SHEET_TS
      };
      row.source_row_hash = buildRowHash(row);
      return row;
    })
    .filter((r) => r.nrp || r.mechanic_name || r.service_resource);
}

/* =========================
   SYNC
========================= */
async function syncSheet(workbook, sheetName, tableName, mapper) {
  const sourceRows = readSheetObjects(workbook, sheetName);
  const payload = mapper(sourceRows);

  console.log(`[${tableName}] source rows: ${sourceRows.length}`);
  await clearTableBySource(tableName, sheetName);
  await insertChunks(tableName, payload);
  console.log(`[${tableName}] synced rows: ${payload.length}`);
}

async function main() {
  try {
    console.log("=== SYNC START ===");
    console.log(`Spreadsheet: ${SPREADSHEET_ID}`);
    console.log(`Tables: ${TABLE_SA}, ${TABLE_TS}`);

    const workbook = await downloadWorkbookFromGoogleSheet();

    await syncSheet(workbook, SHEET_SA, TABLE_SA, mapRasioSaRows);
    await syncSheet(workbook, SHEET_TS, TABLE_TS, mapTimesheetRows);

    console.log("=== SYNC DONE ===");
  } catch (err) {
    console.error("=== SYNC ERROR ===");
    console.error(err.message || err);
    process.exit(1);
  }
}

main();
