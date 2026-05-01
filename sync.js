const XLSX = require("xlsx");

/* =========================
   WAJIB DIISI
========================= */
const SUPABASE_URL = "https://wqgkwbtsfrmpgwxhaybz.supabase.co";
const SUPABASE_SERVICE_ROLE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndxZ2t3YnRzZnJtcGd3eGhheWJ6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzI1ODQ2MzAsImV4cCI6MjA4ODE2MDYzMH0.a4mN97nAiWqHhuhKJGNsmM1qukUdesUCThor-JKmirA";
const SPREADSHEET_ID = "1ijSGUBlsRlbJj_F8iN_RekJbj2obLvEr";

/* =========================
   SESUAIKAN JIKA PERLU
========================= */
const SHEET_SA = "Rasio SA";
const SHEET_TS = "TimeSheet";

const TABLE_SA = "Rasio SA";
const TABLE_TS = "TimeSheet";

const CHUNK_SIZE = 500;

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

function hasAnyValue(row) {
  return Object.values(row).some(
    (v) => v !== null && v !== undefined && String(v).trim() !== ""
  );
}

function buildSupabaseUrl(tableName, params = {}) {
  const url = new URL(`/rest/v1/${encodeURIComponent(tableName)}`, SUPABASE_URL);
  for (const [key, value] of Object.entries(params)) {
    url.searchParams.set(key, value);
  }
  return url.toString();
}

async function supabaseRequest(tableName, method, params = {}, body = null, extraHeaders = {}) {
  const headers = {
    apikey: SUPABASE_SERVICE_ROLE_KEY,
    Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    ...extraHeaders
  };

  if (body !== null) {
    headers["Content-Type"] = "application/json";
  }

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

async function clearTable(tableName, anyColumnName) {
  await supabaseRequest(
    tableName,
    "DELETE",
    {
      or: `(${anyColumnName}.not.is.null,${anyColumnName}.is.null)`
    },
    null,
    {
      Prefer: "return=minimal"
    }
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
      {
        Prefer: "return=minimal"
      }
    );
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
  if (!ws) {
    throw new Error(`Sheet tidak ditemukan: ${sheetName}`);
  }

  return XLSX.utils.sheet_to_json(ws, {
    defval: null,
    raw: false
  }).filter(hasAnyValue);
}

/* =========================
   MAPPING SESUAI HEADER
========================= */
function mapRasioSaRows(rows) {
  return rows
    .map((r) => ({
      "Case_Number": asText(r["Case_Number"]),
      "SA_Number": asText(r["SA_Number"]),
      "Actual_Dispatched_Date": asText(r["Actual_Dispatched_Date"]),
      "Actual_Travel_Date": asText(r["Actual_Travel_Date"]),
      "Actual_InProgress_Date": asText(r["Actual_InProgress_Date"]),
      "Actual_Completed_Date": asText(r["Actual_Completed_Date"]),
      "TimeSheet_Submitted": asText(r["TimeSheet_Submitted"]),
      "TimeSheet_Approved": asText(r["TimeSheet_Approved"]),
      "Case_Description": asText(r["Case_Description"]),
      "Service_Area": asText(r["Service_Area"]),
      "Mechanic_Name": asText(r["Mechanic_Name"]),
      "Scheduled_Start": asText(r["Scheduled_Start"]),
      "SA_Accuracy": asInteger(r["SA_Accuracy"])
    }))
    .filter((r) => r["Case_Number"] || r["SA_Number"] || r["Mechanic_Name"]);
}

function mapTimesheetRows(rows) {
  return rows
    .map((r) => ({
      "Month": asText(r["Month"]),
      "NRP": asText(r["NRP"]),
      "Mechanic_Name": asText(r["Mechanic_Name"]),
      "Available_Hours": asNumber(r["Available_Hours"]),
      "Qty_SA": asNumber(r["Qty_SA"]),
      "ACH": asNumber(r["ACH"]),
      "TDT": asNumber(r["TDT"]),
      "EFH": asNumber(r["EFH"]),
      "JA": asNumber(r["JA"]),
      "JE": asNumber(r["JE"]),
      "UT/Partner": asText(r["UT/Partner"]),
      "Date": asText(r["Date"]),
      "Service_Resource": asText(r["Service_Resource"])
    }))
    .filter((r) => r["NRP"] || r["Mechanic_Name"]);
}

/* =========================
   SYNC
========================= */
async function syncRasioSa(workbook) {
  const sourceRows = readSheetObjects(workbook, SHEET_SA);
  const payload = mapRasioSaRows(sourceRows);

  console.log(`[${TABLE_SA}] source rows: ${sourceRows.length}`);
  await clearTable(TABLE_SA, "Case_Number");
  await insertChunks(TABLE_SA, payload);
  console.log(`[${TABLE_SA}] synced rows: ${payload.length}`);
}

async function syncTimesheet(workbook) {
  const sourceRows = readSheetObjects(workbook, SHEET_TS);
  const payload = mapTimesheetRows(sourceRows);

  console.log(`[${TABLE_TS}] source rows: ${sourceRows.length}`);
  await clearTable(TABLE_TS, "NRP");
  await insertChunks(TABLE_TS, payload);
  console.log(`[${TABLE_TS}] synced rows: ${payload.length}`);
}

async function main() {
  try {
    console.log("=== SYNC START ===");

    const workbook = await downloadWorkbookFromGoogleSheet();

    await syncRasioSa(workbook);
    await syncTimesheet(workbook);

    console.log("=== SYNC DONE ===");
  } catch (err) {
    console.error("=== SYNC ERROR ===");
    console.error(err.message || err);
    process.exit(1);
  }
}

main();