/**
 * INSURANCE FMS DATA PORTAL — Google Apps Script + AWS Athena
 *
 * SETUP
 *  1. Apps Script → Project Settings → Script Properties, add:
 *       AWS_ACCESS_KEY, AWS_SECRET_KEY   (do NOT hard-code them in this file)
 *  2. Reload the Sheet → menu "🚛 Insurance FMS" → "Setup Config Sheet" → "Run Query".
 *
 * "Bytes scanned limit was exceeded" means the workgroup's per-query data
 * limit (see WORK_GROUP) cancelled the query. Changes in this
 * version to scan less data:
 *   - the wide fact table is read exactly once (Athena re-runs a CTE at
 *     every reference); the small dimension tables are joined directly
 *   - redundant GROUP BY / DISTINCT passes removed
 *   - on failure the script reports bytes scanned, so you can see how
 *     close a narrower date range / fewer states gets you to the limit
 * If it still trips, narrow the date range / states in Config, or ask the
 * workgroup admin to raise the per-query limit.
 */

// ---------------- AWS CONFIG ----------------
const AWS_REGION      = 'ap-south-1';
const ATHENA_HOST     = 'athena.' + AWS_REGION + '.amazonaws.com';
const ATHENA_ENDPOINT = 'https://' + ATHENA_HOST + '/';
const WORK_GROUP      = 'analyst-adhoc-executions-highlimit';
const SCHEMA          = 'fact_tables';

const CONFIG_SHEET  = 'Config';
const RESULTS_SHEET = 'Results';

function awsCreds_() {
  const p = PropertiesService.getScriptProperties();
  const key = p.getProperty('AWS_ACCESS_KEY'), secret = p.getProperty('AWS_SECRET_KEY');
  if (!key || !secret) throw new Error('Set AWS_ACCESS_KEY and AWS_SECRET_KEY in Project Settings → Script Properties.');
  return { key: key, secret: secret };
}

// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚛 Insurance FMS')
    .addItem('1️⃣ Setup Config Sheet', 'setupConfigSheet')
    .addItem('▶ Run Query', 'runQuery')
    .addItem('🔄 Fetch Last Result', 'fetchLastResult')
    .addItem('🔍 Diagnose Scan Limit', 'diagnose')
    .addItem('📏 Probe Bytes Per Table', 'probeTables')
    .addItem('📐 Probe Wide Table Columns', 'probeWideColumns')
    .addToUi();
}

// ============================================================
//  CONFIG SHEET
// ============================================================
function setupConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG_SHEET);
  sh.clear();

  const today = new Date();
  const start = new Date(today); start.setDate(start.getDate() - 30);
  const end   = new Date(today); end.setDate(end.getDate() + 45);

  sh.getRange('A1:B6').setValues([
    ['Parameter', 'Value'],
    ['Start Date (expiry from)', Utilities.formatDate(start, 'IST', 'yyyy-MM-dd')],
    ['End Date (expiry to)',     Utilities.formatDate(end,   'IST', 'yyyy-MM-dd')],
    ['State Prefixes (comma separated)', 'MH,GJ'],
    ['Row Limit (0 = no limit)', 0],
    ['Last Query Execution ID', ''],
  ]);
  sh.getRange('A1:B1').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setColumnWidth(1, 260).setColumnWidth(2, 300);
  SpreadsheetApp.getUi().alert('Config sheet ready. Set your dates & states, then run "▶ Run Query".');
}

function readConfig() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!sh) throw new Error('Config sheet not found. Run "Setup Config Sheet" first.');
  const vals = sh.getRange('B2:B5').getValues();

  const fmt = v => (v instanceof Date)
    ? Utilities.formatDate(v, 'IST', 'yyyy-MM-dd')
    : String(v).trim();

  const startDate = fmt(vals[0][0]);
  const endDate   = fmt(vals[1][0]);
  const prefixes  = String(vals[2][0]).split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
  const limit     = parseInt(vals[3][0]) || 0;

  if (!/^\d{4}-\d{2}-\d{2}$/.test(startDate) || !/^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
    throw new Error('Dates must be in yyyy-MM-dd format.');
  }
  if (prefixes.some(p => !/^[A-Z0-9]+$/.test(p))) {
    throw new Error('State prefixes may only contain letters and digits.');
  }
  return { startDate, endDate, prefixes, limit };
}

// ============================================================
//  QUERY BUILDER
// ============================================================
function buildQuery(cfg) {
  const prefixFilter = cfg.prefixes.length
    ? '(' + cfg.prefixes.map(p => "truck_no like '" + p + "%'").join(' or ') + ')'
    : '1=1';
  const limitClause = cfg.limit > 0 ? 'limit ' + cfg.limit : '';

  // The three dimension tables are small (~1 GB combined), so they are joined
  // directly. The wide fact table is the expensive one and is read exactly once.
  return `with
fleet_phone as (
  select id, phone_no
  from supply_team.supply_team_blackbuck_fleetApp_fleetowner
  group by 1, 2
),
tto_latest as (
  select id,
         max_by(truck_id, __ts_ms) as truck_id,
         max_by(kyc_status_v2, __ts_ms) as kyc_status,
         max_by(truck_owner_id, __ts_ms) as truck_owner_id
  from supply_team.supply_team_blackbuck_truck_owner_request
  group by id
),
tto_kyc as (
  select distinct truck_id, truck_owner_id, kyc_status
  from tto_latest
  where kyc_status in ('INSTANT_KYC_APPROVED', 'APPROVED')
),
documents as (
  select id,
         max_by(document_no, version) as document_no,
         max_by(document_type, version) as document_type,
         max_by(entity_id, version) as entity_id,
         max_by(entity_type, version) as entity_type,
         max_by(status, version) as status
  from supply.supply_blackbuck_document
  where document_type in ('PAN_CARD','ADHAAR_CARD') and status = 'APPROVED' and entity_type = 'TRUCK_OWNER'
  group by id
)
select distinct
  m.fleet_owner_id, m.fo_name, m.truck_no, m.truck_id, m.registration_date, m.manufacturer, m.body_type, m.gross_weight,
  m.unladen_weight, m.wheel_base, m.model, m.vehicle_class, m.vehicle_category, m.cubic_capacity,
  m.insurance_name, m.insurance_policy_number, m.insurance_expiry_date, m.financier, m.last_fetched_on,
  p.phone_no, m.latitude, m.longitude, k.kyc_status, k.truck_owner_id,
  b.*, c.document_no as aadhar_no, c.document_type, c.status as aadhar_status
from (
  select fleet_owner_id, fo_name, truck_no, truck_id, registration_date, manufacturer, body_type, gross_weight,
         unladen_weight, wheel_base, model, vehicle_class, vehicle_category, cubic_capacity,
         insurance_name, insurance_policy_number,
         date(from_unixtime((insurance_expiry_date + 19800000)/1000)) as insurance_expiry_date,
         financier, last_updated_on as last_fetched_on, latitude, longitude
  from fact_tables.finserv_insurance_fms_query
  where insurance_expiry_date is not null
    and date(from_unixtime((insurance_expiry_date + 19800000)/1000)) between date('${cfg.startDate}') and date('${cfg.endDate}')
    and ${prefixFilter}
) m
left join fleet_phone p on m.fleet_owner_id = p.id
join tto_kyc k on m.truck_id = k.truck_id
left join (select * from documents where document_type = 'PAN_CARD') b on k.truck_owner_id = b.entity_id
left join (select * from documents where document_type = 'ADHAAR_CARD') c on k.truck_owner_id = c.entity_id
${limitClause}`;
}

// ============================================================
//  MAIN: RUN QUERY
// ============================================================
function runQuery() {
  const ui = SpreadsheetApp.getUi();
  const cfg = readConfig();
  const query = buildQuery(cfg);

  SpreadsheetApp.getActiveSpreadsheet().toast('Submitting query to Athena…', 'Insurance FMS', 10);

  const startResp = athenaCall('StartQueryExecution', {
    QueryString: query,
    QueryExecutionContext: { Database: SCHEMA },
    WorkGroup: WORK_GROUP,
    ClientRequestToken: Utilities.getUuid()
  });
  const execId = startResp.QueryExecutionId;
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET).getRange('B6').setValue(execId);

  const state = pollExecution(execId, 270);

  if (state === 'SUCCEEDED') {
    writeResults(execId);
  } else if (state === 'RUNNING' || state === 'QUEUED') {
    ui.alert('Query is still running on Athena (exceeded Apps Script wait window).\n\n' +
             'Wait a couple of minutes, then use menu → "🔄 Fetch Last Result".');
  } else {
    ui.alert(describeFailure_(execId, state));
  }
}

function describeFailure_(execId, state) {
  const qe = athenaCall('GetQueryExecution', { QueryExecutionId: execId }).QueryExecution;
  const reason = ((qe.Status || {}).StateChangeReason) || 'Unknown';
  const bytes = ((qe.Statistics || {}).DataScannedInBytes);
  let msg = 'Query ' + state + ':\n' + reason;
  if (bytes !== undefined) msg += '\n\nData scanned before stopping: ' + (bytes / 1073741824).toFixed(2) + ' GB';
  if (/bytes scanned/i.test(reason)) {
    msg += '\n\nThe workgroup\'s per-query scan limit was hit. Narrow the date range or ' +
           'state prefixes in Config and retry, or ask the workgroup admin to raise the limit.';
  }
  return msg;
}

function fetchLastResult() {
  const ui = SpreadsheetApp.getUi();
  const execId = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG_SHEET).getRange('B6').getValue();
  if (!execId) { ui.alert('No previous execution found. Run a query first.'); return; }

  const state = athenaCall('GetQueryExecution', { QueryExecutionId: execId }).QueryExecution.Status.State;
  if (state === 'SUCCEEDED') {
    writeResults(execId);
  } else if (state === 'RUNNING' || state === 'QUEUED') {
    ui.alert('Query is still ' + state + '. Try again in a minute.');
  } else {
    ui.alert(describeFailure_(execId, state));
  }
}

function pollExecution(execId, maxSeconds) {
  const deadline = Date.now() + maxSeconds * 1000;
  let state = 'QUEUED';
  while (Date.now() < deadline) {
    const detail = athenaCall('GetQueryExecution', { QueryExecutionId: execId });
    state = detail.QueryExecution.Status.State;
    if (state === 'SUCCEEDED' || state === 'FAILED' || state === 'CANCELLED') return state;
    Utilities.sleep(5000);
  }
  return state;
}

// ============================================================
//  RESULTS → SHEET
// ============================================================
function writeResults(execId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RESULTS_SHEET);
  if (!sh) sh = ss.insertSheet(RESULTS_SHEET);
  sh.clear();

  let nextToken = null;
  let header = null;
  const allRows = [];
  let firstPage = true;

  do {
    const params = { QueryExecutionId: execId, MaxResults: 1000 };
    if (nextToken) params.NextToken = nextToken;
    const resp = athenaCall('GetQueryResults', params);

    const rows = resp.ResultSet.Rows || [];
    let startIdx = 0;
    if (firstPage) {
      header = rows[0].Data.map(d => d.VarCharValue || '');
      startIdx = 1;
      firstPage = false;
    }
    for (let i = startIdx; i < rows.length; i++) {
      allRows.push(rows[i].Data.map(d => (d && d.VarCharValue !== undefined) ? d.VarCharValue : ''));
    }
    nextToken = resp.NextToken || null;
    if (allRows.length > 190000) break;
  } while (nextToken);

  if (!header) { SpreadsheetApp.getUi().alert('No results returned.'); return; }

  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  if (allRows.length) {
    sh.getRange(2, 1, allRows.length, header.length).setValues(allRows);
  }
  sh.setFrozenRows(1);
  ss.toast('Done: ' + allRows.length + ' rows written to "' + RESULTS_SHEET + '"', 'Insurance FMS', 8);
}

// ============================================================
//  AWS SIGNATURE V4  (Athena JSON API)
// ============================================================
function athenaCall(action, payloadObj) {
  const creds = awsCreds_();
  const payload = JSON.stringify(payloadObj);
  const now = new Date();
  const amzDate   = Utilities.formatDate(now, 'GMT', "yyyyMMdd'T'HHmmss'Z'");
  const dateStamp = Utilities.formatDate(now, 'GMT', 'yyyyMMdd');

  const service = 'athena';
  const target  = 'AmazonAthena.' + action;
  const contentType = 'application/x-amz-json-1.1';

  const canonicalHeaders =
    'content-type:' + contentType + '\n' +
    'host:' + ATHENA_HOST + '\n' +
    'x-amz-date:' + amzDate + '\n' +
    'x-amz-target:' + target + '\n';
  const signedHeaders = 'content-type;host;x-amz-date;x-amz-target';
  const canonicalRequest = ['POST', '/', '', canonicalHeaders, signedHeaders, sha256Hex(payload)].join('\n');

  const credentialScope = dateStamp + '/' + AWS_REGION + '/' + service + '/aws4_request';
  const stringToSign = ['AWS4-HMAC-SHA256', amzDate, credentialScope, sha256Hex(canonicalRequest)].join('\n');

  const kDate    = hmac(dateStamp, 'AWS4' + creds.secret);
  const kRegion  = hmacBytes(AWS_REGION, kDate);
  const kService = hmacBytes(service, kRegion);
  const kSigning = hmacBytes('aws4_request', kService);
  const signature = bytesToHex(hmacBytes(stringToSign, kSigning));

  const authHeader = 'AWS4-HMAC-SHA256 Credential=' + creds.key + '/' + credentialScope +
    ', SignedHeaders=' + signedHeaders + ', Signature=' + signature;

  const resp = UrlFetchApp.fetch(ATHENA_ENDPOINT, {
    method: 'post',
    contentType: contentType,
    payload: payload,
    headers: { 'X-Amz-Date': amzDate, 'X-Amz-Target': target, 'Authorization': authHeader },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) throw new Error('Athena ' + action + ' failed (' + code + '): ' + body);
  return JSON.parse(body);
}

function sha256Hex(str) {
  return bytesToHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8));
}
function hmac(message, key) {
  return Utilities.computeHmacSha256Signature(Utilities.newBlob(message).getBytes(), Utilities.newBlob(key).getBytes());
}
function hmacBytes(message, keyBytes) {
  return Utilities.computeHmacSha256Signature(Utilities.newBlob(message).getBytes(), keyBytes);
}
function bytesToHex(bytes) {
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

// ============================================================
//  DIAGNOSE: workgroup limit + table partitioning (scans 0 bytes)
// ============================================================
function diagnose() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Diagnose') || ss.insertSheet('Diagnose');
  sh.clear();
  const rows = [['Item', 'Detail']];

  const wg = athenaCall('GetWorkGroup', { WorkGroup: WORK_GROUP }).WorkGroup.Configuration || {};
  const cutoff = wg.BytesScannedCutoffPerQuery;
  rows.push(['Per-query bytes limit', cutoff ? (cutoff / 1073741824).toFixed(2) + ' GB' : 'none reported']);

  const tables = [
    'fact_tables.finserv_insurance_fms_query',
    'supply_team.supply_team_blackbuck_fleetApp_fleetowner',
    'supply_team.supply_team_blackbuck_truck_owner_request',
    'supply.supply_blackbuck_document'
  ];
  tables.forEach(t => {
    try {
      const id = athenaCall('StartQueryExecution', {
        QueryString: 'SHOW CREATE TABLE ' + t, WorkGroup: WORK_GROUP,
        ClientRequestToken: Utilities.getUuid()
      }).QueryExecutionId;
      const st = pollExecution(id, 60);
      if (st !== 'SUCCEEDED') {
        const why = (athenaCall('GetQueryExecution', { QueryExecutionId: id }).QueryExecution.Status || {}).StateChangeReason;
        rows.push([t, 'Query ' + st + ': ' + why]); return;
      }
      const res = athenaCall('GetQueryResults', { QueryExecutionId: id, MaxResults: 1000 });
      const ddl = res.ResultSet.Rows.map(r => (r.Data[0] || {}).VarCharValue || '').join('\n');
      rows.push([t, ddl]);
    } catch (e) { rows.push([t, 'Error: ' + e.message]); }
  });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange('A1:B1').setFontWeight('bold');
  sh.setColumnWidth(1, 300).setColumnWidth(2, 700);
  ss.toast('Diagnose written to "Diagnose" sheet', 'Insurance FMS', 8);
}

// ============================================================
//  PROBE: bytes each table costs for the columns the query reads
//  (each probe is capped at the workgroup limit; "50.00 GB / CANCELLED"
//   means that table alone needs at least that much)
// ============================================================
function probeTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = readConfig();
  const probes = [
    ['finserv_insurance_fms_query (filtered, as in query)',
     "select count(*) from fact_tables.finserv_insurance_fms_query where insurance_expiry_date is not null " +
     "and date(from_unixtime((insurance_expiry_date + 19800000)/1000)) between date('" + cfg.startDate + "') and date('" + cfg.endDate + "')"],
    ['fleetowner (id, phone_no)',
     'select count(id), count(phone_no) from supply_team.supply_team_blackbuck_fleetApp_fleetowner'],
    ['truck_owner_request (id, truck_id, kyc_status_v2, truck_owner_id, __ts_ms)',
     'select count(id), count(truck_id), count(kyc_status_v2), count(truck_owner_id), count(__ts_ms) from supply_team.supply_team_blackbuck_truck_owner_request'],
    ['document (id, document_no, document_type, entity_id, entity_type, status, version)',
     'select count(id), count(document_no), count(document_type), count(entity_id), count(entity_type), count(status), count(version) from supply.supply_blackbuck_document']
  ];
  let sh = ss.getSheetByName('Probe') || ss.insertSheet('Probe');
  sh.clear();
  const rows = [['Table / columns', 'State', 'GB scanned', 'Note']];
  probes.forEach(p => {
    try {
      const id = athenaCall('StartQueryExecution', {
        QueryString: p[1], WorkGroup: WORK_GROUP, QueryExecutionContext: { Database: SCHEMA },
        ClientRequestToken: Utilities.getUuid()
      }).QueryExecutionId;
      const st = pollExecution(id, 150);
      const qe = athenaCall('GetQueryExecution', { QueryExecutionId: id }).QueryExecution;
      const gb = ((qe.Statistics || {}).DataScannedInBytes || 0) / 1073741824;
      rows.push([p[0], st, gb.toFixed(2), (qe.Status || {}).StateChangeReason || '']);
    } catch (e) { rows.push([p[0], 'ERROR', '', e.message]); }
    sh.getRange(1, 1, rows.length, 4).setValues(rows);
    SpreadsheetApp.flush();
  });
  sh.getRange('A1:D1').setFontWeight('bold');
  sh.setColumnWidth(1, 520).setColumnWidth(4, 400);
  ss.toast('Probe written to "Probe" sheet', 'Insurance FMS', 8);
}

// ============================================================
//  PROBE: GB scanned per column of the wide fact table
//  (all probes start together, then are polled, to fit the 6-min cap)
// ============================================================
function probeWideColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = readConfig();
  const cols = ['fleet_owner_id', 'fo_name', 'truck_no', 'truck_id', 'registration_date', 'manufacturer', 'body_type',
    'gross_weight', 'unladen_weight', 'wheel_base', 'model', 'vehicle_class', 'vehicle_category', 'cubic_capacity',
    'insurance_name', 'insurance_policy_number', 'insurance_expiry_date', 'financier', 'last_updated_on', 'latitude', 'longitude'];
  const where = " where insurance_expiry_date is not null and date(from_unixtime((insurance_expiry_date + 19800000)/1000)) between date('" +
    cfg.startDate + "') and date('" + cfg.endDate + "')";

  const jobs = cols.map(c => ({
    col: c,
    id: athenaCall('StartQueryExecution', {
      QueryString: 'select count(' + c + ') from fact_tables.finserv_insurance_fms_query' + where,
      WorkGroup: WORK_GROUP, QueryExecutionContext: { Database: SCHEMA },
      ClientRequestToken: Utilities.getUuid()
    }).QueryExecutionId
  }));

  const deadline = Date.now() + 270 * 1000;
  const rows = [['Column', 'State', 'GB scanned']];
  let pending = jobs.slice(), done = [];
  while (pending.length && Date.now() < deadline) {
    const still = [];
    pending.forEach(j => {
      const qe = athenaCall('GetQueryExecution', { QueryExecutionId: j.id }).QueryExecution;
      const st = qe.Status.State;
      if (st === 'QUEUED' || st === 'RUNNING') { still.push(j); return; }
      done.push([j.col, st, (((qe.Statistics || {}).DataScannedInBytes || 0) / 1073741824)]);
    });
    pending = still;
    if (pending.length) Utilities.sleep(5000);
  }
  pending.forEach(j => done.push([j.col, 'STILL RUNNING', 0]));
  let total = 0;
  done.sort((a, b) => b[2] - a[2]).forEach(r => { total += r[2]; rows.push([r[0], r[1], r[2].toFixed(2)]); });
  rows.push(['TOTAL (sum of columns)', '', total.toFixed(2)]);

  let sh = ss.getSheetByName('ProbeColumns') || ss.insertSheet('ProbeColumns');
  sh.clear();
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange('A1:C1').setFontWeight('bold');
  sh.setColumnWidth(1, 260);
  ss.toast('Wrote "ProbeColumns" sheet', 'Insurance FMS', 8);
}
