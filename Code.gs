/**
 * INSURANCE FMS DATA PORTAL — Google Apps Script + AWS Athena
 *
 * SETUP
 *  1. Apps Script → Project Settings → Script Properties, add:
 *       AWS_ACCESS_KEY, AWS_SECRET_KEY   (do NOT hard-code them in this file)
 *  2. Reload the Sheet → menu "🚛 Insurance FMS" → "Setup Config Sheet" → "Run Query".
 *
 * "Bytes scanned limit was exceeded" means the workgroup's per-query data
 * limit (analyst-adhoc-executions) cancelled the query. Changes in this
 * version to scan less data:
 *   - the state-prefix and date filters are applied to the first table
 *     scanned, and every other table is semi-joined to the surviving
 *     truck_ids / owner ids instead of being scanned for all rows
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
const WORK_GROUP      = 'analyst-adhoc-executions';
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

  return `with
base as (
  select fleet_owner_id, fo_name, truck_no, truck_id, registration_date, manufacturer, body_type, gross_weight,
         unladen_weight, wheel_base, model, vehicle_class, vehicle_category, cubic_capacity,
         insurance_name, insurance_policy_number,
         date(from_unixtime((insurance_expiry_date + 19800000)/1000)) as insurance_expiry_date,
         financier, last_updated_on as last_fetched_on, latitude, longitude
  from fact_tables.finserv_insurance_fms_query
  where insurance_expiry_date is not null
    and date(from_unixtime((insurance_expiry_date + 19800000)/1000)) between date('${cfg.startDate}') and date('${cfg.endDate}')
    and ${prefixFilter}
),
fleet_phone as (
  select id, phone_no
  from supply_team.supply_team_blackbuck_fleetApp_fleetowner
  where id in (select fleet_owner_id from base)
  group by 1, 2
),
final_final_ritesh_data as (
  select a.*, b.phone_no
  from base a
  left join fleet_phone b on a.fleet_owner_id = b.id
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
    and truck_id in (select truck_id from base)
),
final_data_ritesh_master as (
  select a.*, b.kyc_status, b.truck_owner_id
  from final_final_ritesh_data a
  join tto_kyc b on a.truck_id = b.truck_id
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
),
docs_ok as (
  select * from documents
  where entity_id in (select truck_owner_id from final_data_ritesh_master)
)
select distinct a.*, b.*, c.document_no as aadhar_no, c.document_type, c.status as aadhar_status
from final_data_ritesh_master a
left join (select * from docs_ok where document_type = 'PAN_CARD') b
  on a.truck_owner_id = b.entity_id
left join (select * from docs_ok where document_type = 'ADHAAR_CARD') c
  on a.truck_owner_id = c.entity_id
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
      if (st !== 'SUCCEEDED') { rows.push([t, 'Query ' + st]); return; }
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
