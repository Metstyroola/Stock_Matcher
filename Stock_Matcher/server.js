/**
 * StockMatch Server v5
 * Fixed: gcloud PATH, token caching, NDJSON paging
 */
'use strict';

const http   = require('http');
const https  = require('https');
const fs     = require('fs');
const path   = require('path');
const crypto = require('crypto');
const cp     = require('child_process');
const rl_mod = require('readline');

const PORT       = process.env.PORT || 3000;
const PROJECT_ID = 'heroic-ruler-198603';
const DATASET    = 'Sales_Dashboard_Views';
const TABLE      = 'suppliers_check_dashboard_table';
const NZ_DATASET = 'Sales_Dasboard_Views_NZ';   // note: typo in BQ ("Dasboard") is intentional
const NZ_TABLE   = 'suppliers_check_report_table';
const APP_FILE   = path.join(__dirname, 'stock_matcher_app.html');
const ENV_FILE   = path.join(__dirname, '.env');

// Load .env
if (fs.existsSync(ENV_FILE)) {
  fs.readFileSync(ENV_FILE,'utf8').split('\n').forEach(line=>{
    const m=line.match(/^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.+?)\s*$/);
    if(m&&!process.env[m[1]])process.env[m[1]]=m[2].replace(/^['"]|['"]$/g,'');
  });
}

// ── Auto-install xlsx package if missing ──────────────────────────────────
let XLSX_LIB = null;
function getXLSX() {
  if (XLSX_LIB) return XLSX_LIB;
  // Try locations in order
  const tries = [
    ()=>require('xlsx'),
    ()=>require(path.join(__dirname,'node_modules','xlsx','lib','xlsx.js')),
    ()=>require(path.join(__dirname,'node_modules','xlsx','xlsx.js')),
  ];
  for (const fn of tries) {
    try { XLSX_LIB = fn(); return XLSX_LIB; } catch(e) {}
  }
  // Auto-install
  try {
    console.log('[XLSX] Installing xlsx package (first run)...');
    cp.execSync('npm install xlsx --save --no-audit --no-fund', {
      cwd: __dirname, timeout: 120000, stdio: 'pipe'
    });
    XLSX_LIB = require(path.join(__dirname,'node_modules','xlsx','lib','xlsx.js'));
    console.log('[XLSX] Installed successfully');
    return XLSX_LIB;
  } catch(e) {
    console.error('[XLSX] Install failed:', e.message);
    return null;
  }
}

// ── gcloud executable path ─────────────────────────────────────────────────
// Node.js spawned processes don't inherit the full user PATH on Windows
// Find gcloud in all known install locations
function findGcloud() {
  const candidates = [
    // From environment
    process.env.GCLOUD_PATH,
    // Standard Windows install locations
    path.join(process.env.LOCALAPPDATA||'', 'Google', 'Cloud SDK', 'google-cloud-sdk', 'bin', 'gcloud.cmd'),
    path.join(process.env.PROGRAMFILES||'', 'Google', 'Cloud SDK', 'google-cloud-sdk', 'bin', 'gcloud.cmd'),
    path.join(process.env['PROGRAMFILES(X86)']||'', 'Google', 'Cloud SDK', 'google-cloud-sdk', 'bin', 'gcloud.cmd'),
    path.join(process.env.USERPROFILE||'', 'AppData', 'Local', 'Google', 'Cloud SDK', 'google-cloud-sdk', 'bin', 'gcloud.cmd'),
    // In PATH
    'gcloud',
    'gcloud.cmd',
  ].filter(Boolean);

  for (const c of candidates) {
    try {
      if (c === 'gcloud' || c === 'gcloud.cmd') {
        // Check if it's in PATH
        cp.execSync(`"${c}" --version`, { stdio: 'ignore', timeout: 5000 });
        return c;
      }
      if (fs.existsSync(c)) return c;
    } catch(e) {}
  }
  return null;
}

let _gcloudPath = null;
function getGcloud() {
  if (_gcloudPath) return _gcloudPath;
  _gcloudPath = findGcloud();
  if (_gcloudPath) console.log('[Auth] gcloud found:', _gcloudPath);
  else console.log('[Auth] gcloud not found in PATH or standard locations');
  return _gcloudPath;
}

// ── HTTP helpers ───────────────────────────────────────────────────────────
function post(hostname,urlPath,body,headers={}) {
  return new Promise((resolve,reject)=>{
    const buf=typeof body==='string'?Buffer.from(body):body;
    const req=https.request({hostname,path:urlPath,method:'POST',
      headers:{'Content-Length':buf.length,...headers}},res=>{
      const chunks=[];
      res.on('data',c=>chunks.push(c));
      res.on('end',()=>{
        const d=Buffer.concat(chunks).toString('utf8');
        try{resolve(JSON.parse(d));}
        catch(e){
          console.error(`[HTTP] POST ${hostname}${urlPath} → HTTP ${res.statusCode}`);
          console.error(`[HTTP] Response body: ${d.slice(0,400)}`);
          if(res.statusCode===404){
            // 404 from BQ API = token lacks bigquery scope
            // Run Setup_BigQuery_Auth.bat to fix
            reject(new Error('NOT_AUTHENTICATED'));
          } else {
            reject(new Error(`HTTP ${res.statusCode} from BigQuery API`));
          }
        }
      });
    });
    req.on('error',reject);req.write(buf);req.end();
  });
}

function get(hostname,urlPath,headers={}) {
  return new Promise((resolve,reject)=>{
    const req=https.request({hostname,path:urlPath,method:'GET',headers},res=>{
      const chunks=[];
      res.on('data',c=>chunks.push(c));
      res.on('end',()=>{
        const d=Buffer.concat(chunks).toString('utf8');
        // Log every BQ API response status
        if(hostname.includes('bigquery')){
          console.log(`[HTTP GET ${res.statusCode}] ${urlPath.slice(0,100)}`);
          if(res.statusCode!==200) console.log(`[HTTP BODY] ${d.slice(0,500)}`);
        }
        // Always try to parse JSON — even error responses from BQ are JSON
        try {
          resolve(JSON.parse(d));
        } catch(e){
          console.error(`[HTTP GET] ${hostname}${urlPath.slice(0,80)} → HTTP ${res.statusCode}: ${d.slice(0,300)}`);
          resolve({error:{code:res.statusCode,message:`HTTP ${res.statusCode}: ${d.slice(0,200)}`}});
        }
      });
    });
    req.on('error',e=>{
      console.error(`[HTTP GET] ${hostname}${urlPath.slice(0,80)} → network error: ${e.message}`);
      reject(e);
    });
    req.end();
  });
}

function httpGet(hostname,urlPath,headers={}) {
  return new Promise((resolve,reject)=>{
    const req=require('http').request({hostname,path:urlPath,method:'GET',headers},res=>{
      const chunks=[];
      res.on('data',c=>chunks.push(c));
      res.on('end',()=>{
        const d=Buffer.concat(chunks).toString('utf8');
        try{resolve({status:res.statusCode,body:JSON.parse(d)});}
        catch(e){resolve({status:res.statusCode,body:d});}
      });
    });
    req.on('error',reject);req.end();
  });
}

// ── Token ──────────────────────────────────────────────────────────────────
let _cachedToken=null, _tokenExpiry=0;

function runCmd(cmd) {
  return new Promise((res,rej)=>
    cp.exec(cmd, {timeout:15000, windowsHide:true}, (err,out,stderr)=>
      err ? rej(new Error(stderr||err.message)) : res(out.trim())
    )
  );
}

const BQ_SCOPE = 'https://www.googleapis.com/auth/bigquery https://www.googleapis.com/auth/cloud-platform';

async function refreshUserToken(c) {
  // Always request with explicit BQ scope — fixes HTTP 404 from BQ API
  const body = [
    `client_id=${encodeURIComponent(c.client_id)}`,
    `client_secret=${encodeURIComponent(c.client_secret)}`,
    `refresh_token=${encodeURIComponent(c.refresh_token)}`,
    `grant_type=refresh_token`,
    `scope=${encodeURIComponent(BQ_SCOPE)}`
  ].join('&');
  const r = await post('oauth2.googleapis.com', '/token', body,
    {'Content-Type':'application/x-www-form-urlencoded'});
  if (r.error) {
    if (r.error==='invalid_rapt'||r.error==='reauth_required')
      throw new Error('NOT_AUTHENTICATED');
    throw new Error('Token refresh: '+(r.error_description||r.error));
  }
  if (!r.access_token) throw new Error('No access_token in refresh response');
  return r.access_token;
}

async function tokenHasBQScope(token) {
  try {
    const info = await get('oauth2.googleapis.com',
      `/tokeninfo?access_token=${encodeURIComponent(token)}`);
    const scope = info.scope || '';
    const ok = scope.includes('bigquery') || scope.includes('cloud-platform');
    if (!ok) console.log('[Auth] Token lacks BQ scope. Has:', scope.slice(0,200));
    return ok;
  } catch(e) { return true; } // assume OK if check fails
}

async function getADCToken() {
  if (_cachedToken && Date.now() < _tokenExpiry - 60000) return _cachedToken;
  _cachedToken = null;
  const set = t => { _cachedToken=t; _tokenExpiry=Date.now()+3500000; return t; };
  const gc = getGcloud();

  // ── Method 1: gcloud user token (ALWAYS works, uses verified gcloud OAuth client) ──
  // This bypasses the "sensitive info" block from ADC unverified app
  if (gc) {
    try {
      const t = await runCmd(`"${gc}" auth print-access-token`);
      if (t && t.length > 20 && !t.includes('\n') && !t.includes('ERROR')) {
        console.log('[Auth] ✓ gcloud user token');
        return set(t);
      }
    } catch(e) { console.log('[Auth] gcloud user token:', e.message); }
  }

  // ── Method 2: gcloud ADC token ─────────────────────────────────────────────
  if (gc) {
    try {
      const t = await runCmd(`"${gc}" auth application-default print-access-token`);
      if (t && t.length > 20 && !t.includes('\n') && !t.includes('ERROR')) {
        console.log('[Auth] ✓ gcloud ADC token');
        return set(t);
      }
    } catch(e) { console.log('[Auth] gcloud ADC token:', e.message); }
  }

  // ── Method 3: Service account key ─────────────────────────────────────────
  const saPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  if (saPath && fs.existsSync(saPath)) {
    try {
      const sa = JSON.parse(fs.readFileSync(saPath,'utf8'));
      if (sa.type === 'service_account') {
        console.log('[Auth] ✓ service account key');
        return set(await getSAToken(sa));
      }
    } catch(e) { console.log('[Auth] SA key:', e.message); }
  }

  // ── Method 4: GCE metadata ─────────────────────────────────────────────────
  try {
    const r = await httpGet('metadata.google.internal',
      '/computeMetadata/v1/instance/service-accounts/default/token',
      {'Metadata-Flavor':'Google'});
    if (r.status===200 && r.body?.access_token) {
      console.log('[Auth] ✓ GCE metadata token');
      _cachedToken=r.body.access_token;
      _tokenExpiry=Date.now()+(r.body.expires_in||3500)*1000;
      return _cachedToken;
    }
  } catch(e) {}

  throw new Error('NOT_AUTHENTICATED');
}


function b64url(b){return b.toString('base64').replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');}
let _saTok=null,_saExp=0;
async function getSAToken(sa) {
  if (_saTok&&Date.now()<_saExp) return _saTok;
  const now=Math.floor(Date.now()/1000);
  const h=b64url(Buffer.from(JSON.stringify({alg:'RS256',typ:'JWT'})));
  const c=b64url(Buffer.from(JSON.stringify({
    iss:sa.client_email, scope:BQ_SCOPE,
    aud:'https://oauth2.googleapis.com/token', exp:now+3600, iat:now})));
  const u=`${h}.${c}`;
  const sig=crypto.createSign('RSA-SHA256').update(u).sign(sa.private_key);
  const jwt=`${u}.${b64url(sig)}`;
  const r=await post('oauth2.googleapis.com','/token',
    `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
    {'Content-Type':'application/x-www-form-urlencoded'});
  if (r.error) throw new Error('SA auth: '+r.error_description);
  _saTok=r.access_token; _saExp=Date.now()+3500000; return _saTok;
}


// ── BigQuery ───────────────────────────────────────────────────────────────
const SC_SUPPLIERS = [
  'atlastyres','bsatyres','chapelcorner','continental','cypresstyres',
  'freedomtyres','goodride','goodride-wa','logictyres','michelin',
  'neta','newbee','onyx','pirelli','spmotorcycles','stgeorge','superior',
  'tempetyres','toptyres','townsend','tyredepot','tyrenetwork',
  'tyreprofessionals','yhi','yokohama'
];

const SC_SQL = `
  SELECT supplier_pid, spinach_pid AS spinach_id, data_provider_pid,
    manufacturer, profile, profile_text, dimensions,
    name_supplier AS supplier_name, \`key\` AS supplier_key
  FROM \`${PROJECT_ID}.${DATASET}.${TABLE}\`
  WHERE is_active = true AND is_active_stock = true
    AND \`key\` IN (${SC_SUPPLIERS.map(s=>`'${s}'`).join(',')})`;

const NZ_SC_SQL = `
  SELECT supplier_pid, spinach_pid AS spinach_id, data_provider_pid,
    manufacturer, profile, profile_text, dimensions,
    name_supplier AS supplier_name, \`key\` AS supplier_key
  FROM \`${PROJECT_ID}.${NZ_DATASET}.${NZ_TABLE}\`
  WHERE is_active = true AND is_active_stock = true`;

// ── Stock-mapping dashboard (live supplier_stock) ────────────────────────────
// Powers supplier_stock_match_dashboard_live.html. Reproduces the curated
// snapshot's methodology: qty>0, exclude DELETED pids, and restrict to the
// hand-curated 31-supplier account set (via the supplier_name -> key allow-list
// below, extracted from the original static dashboard's own location lists).
const STOCK_STOCK_TABLE = 'xx_development.popeye_production_au_public_supplier_stock';
const STOCK_MAP_NAME_TO_KEY = {
  "Tempetyres VIC": "tempetyres", "Tempetyres QLD": "tempetyres", "Tempetyres NSW": "tempetyres",
  "Tempetyres": "tempetyres", "Tempetyres SA": "tempetyres", "Tempetyres WA": "tempetyres",
  "Tempetyres Special NSW": "tempetyres", "Tempetyres Special VIC": "tempetyres",
  "Tempetyres Special SA": "tempetyres", "Tempetyres Special WA": "tempetyres",
  "Total Tyres Virtual VIC": "totaltyres", "Total Tyres Virtual NSW": "totaltyres",
  "Total Tyres Virtual WA": "totaltyres", "Total Tyres Virtual QLD": "totaltyres",
  "Total Tyres Virtual SA": "totaltyres", "Total Tyres VIC": "totaltyres", "Total Tyres WA": "totaltyres",
  "Total Tyres NSW": "totaltyres", "Total Tyres TOW": "totaltyres", "Total Tyres TAS": "totaltyres",
  "Total Tyres SA": "totaltyres",
  "Yhi NSW": "yhi", "Yhi QLD": "yhi", "Yhi VIC": "yhi", "Yhi WA": "yhi", "Yhi SA": "yhi", "Yhi TAS": "yhi",
  "Newbee NSW Tyres Virtual": "newbee", "Newbee NSW": "newbee", "Newbee QLD Tyres": "newbee",
  "Newbee VIC Tyres": "newbee", "Newbee Wollongong Tyres": "newbee",
  "Goodride NSW Riverwood": "goodride", "Goodride": "goodride", "Goodride NSW Padstow": "goodride",
  "Goodride NSW Gosford": "goodride", "Goodride ACT": "goodride", "Goodride NSW Wollongong": "goodride",
  "Yokohama NSW": "yokohama", "Yokohama VIC": "yokohama", "Yokohama QLD Brisbane": "yokohama",
  "Yokohama WA": "yokohama", "Yokohama NT": "yokohama", "Yokohama QLD Townsville": "yokohama",
  "Yokohama SA": "yokohama",
  "Neta Tyres Brisbane Qld": "neta", "Neta Tyres Gold Coast Qld": "neta", "Neta Tyres Rockhampton Qld": "neta",
  "Neta Tyres Cairns Qld": "neta", "Neta Tyres Toowoomba Qld": "neta", "Neta Tyres Sunshine Coast Qld": "neta",
  "Neta Tyres Lismore Nsw": "neta",
  "Pirelli NSW": "pirelli", "Pirelli VIC": "pirelli", "Pirelli QLD": "pirelli", "Pirelli WA": "pirelli",
  "Tyre Connect": "tyreconnect",
  "Hankook WA": "hankook", "Hankook NSW": "hankook", "Hankook VIC": "hankook", "Hankook QLD": "hankook",
  "Hankook SA": "hankook",
  "Chapel Corner RPM": "chapelcorner", "Chapel Corner Goodyear": "chapelcorner",
  "Chapel Corner Virtual": "chapelcorner", "Chapel Corner Megabus": "chapelcorner",
  "Chapel Corner MPS": "chapelcorner",
  "Logic Kumho WA": "logictyres", "Logic Kumho VIC": "logictyres", "Logic Kumho QLD": "logictyres",
  "Logic Kumho NSW Marsden Park": "logictyres", "Logic Kumho SA": "logictyres",
  "Logic Tyres NSW Seven Hills": "logictyres",
  "Continental Tyres WA": "continental", "Continental Tyres QLD": "continental",
  "Continental Tyres VIC": "continental", "Continental Tyres NSW": "continental",
  "Toptyres QLD": "toptyres", "Toptyres NSW": "toptyres", "Toptyres TOW": "toptyres", "Toptyres SA": "toptyres",
  "Cypress Tyres Qld Brisbane": "cypresstyres", "Cypress Tyres Nsw": "cypresstyres",
  "Cypress Tyres Qld Gold Coast": "cypresstyres", "Cypress Tyres VIC Pty Ltd": "cypresstyres",
  "The Tyre Professionals Virtual": "tyreprofessionals", "The Tyre Professionals": "tyreprofessionals",
  "Freedom Tyres VIC": "freedomtyres", "Freedom Tyres QLD": "freedomtyres",
  "Freedom Tyres NSW": "freedomtyres", "Freedom Tyres WA": "freedomtyres",
  "Onyx Vic": "onyx", "Onyx Tyres Qld Brisbane": "onyx", "OLD Onyx Tyres NSW": "onyx",
  "OLD Onyx Tyres Qld Gold Coast": "onyx", "Onyx Qld Sunshine Coast": "onyx",
  "Tyregenie": "tyregenie",
  "St George Tyres": "stgeorge",
  "Tyrenetwork VIC": "tyrenetwork", "Tyrenetwork QLD": "tyrenetwork", "Tyrenetwork": "tyrenetwork",
  "Tyrenetwork NSW": "tyrenetwork", "Tyrenetwork SA": "tyrenetwork",
  "Superior Tyres": "superior",
  "BSA Tyres NSW": "bsatyres",
  "Kumho": "kumho",
  "Tyredepot": "tyredepot", "Tyredepot Virtual": "tyredepot",
  "Townsend": "townsend",
  "GetAGrip": "getagrip",
  "SPMotorcycles": "spmotorcycles",
  "Goodride WA": "goodride-wa",
  "Atlastyres": "atlastyres",
  "Michelin VIC": "michelin", "Michelin NSW": "michelin", "Michelin WA": "michelin", "Michelin QLD": "michelin"
};
const STOCK_ALLOWED_NAMES = Object.keys(STOCK_MAP_NAME_TO_KEY);
const STOCK_KNOWN_KEYS = Array.from(new Set(Object.values(STOCK_MAP_NAME_TO_KEY)));

function sqlQuoteList(strs) {
  return strs.map(s => "'" + String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'").join(',');
}

// Row-level export (CSV/Excel): one row per stock item, for the selected
// supplier key(s) + match_status filter. `keys` must already be validated
// against STOCK_KNOWN_KEYS by the caller (route handler).
function stockRowsForKeys(keys) {
  const keySet = new Set(keys);
  return STOCK_ALLOWED_NAMES.filter(name => keySet.has(STOCK_MAP_NAME_TO_KEY[name]));
}

function stockMappingRowsSQL(names, status) {
  let sql = 'SELECT supplier_name, supplier_pid, supplier_description, qty, price, match_status\n' +
    'FROM `' + PROJECT_ID + '.' + STOCK_STOCK_TABLE + '`\n' +
    "WHERE qty > 0 AND supplier_pid NOT LIKE r'DELETED\\_%'\n" +
    '  AND supplier_name IN (' + sqlQuoteList(names) + ')\n';
  if (status === 'matched' || status === 'unmatched') {
    sql += "  AND match_status = '" + status + "'\n";
  }
  sql += 'ORDER BY supplier_name, supplier_pid';
  return sql;
}

async function fetchStockMappingRows(keys, status) {
  const names = stockRowsForKeys(keys);
  if (!names.length) return [];
  const rows = await runBQQuery(stockMappingRowsSQL(names, status));
  return rows.map(r => ({
    supplier_key: STOCK_MAP_NAME_TO_KEY[r.supplier_name] || '',
    supplier_name: r.supplier_name || '',
    supplier_pid: r.supplier_pid || '',
    description: r.supplier_description || '',
    qty: Number(r.qty) || 0,
    price: (r.price === null || r.price === undefined || r.price === '') ? null : Number(r.price),
    match_status: r.match_status || ''
  }));
}

// Wheel-vs-tyre classification for a supplier's stock descriptions. Validated
// 2026-07-15 against Tempetyres: the CB (center bore) + ET (offset) dimension
// notation catches genuine aftermarket alloy wheel listings that don't even
// contain the word "wheel" (e.g. "20x10.0 5x120 ET48 CB73.1 Satin Black").
const WHEEL_PATTERN_SQL =
  "REGEXP_CONTAINS(UPPER(supplier_description), r'\\bWHEEL[S]?\\b|\\bRIM[S]?\\b|\\bALLOY\\b|\\bWHL\\b|\\bHUB\\b|\\bSPOKE[S]?\\b|\\bSTEELIE[S]?\\b|\\bMAG\\b')\n" +
  "    OR REGEXP_CONTAINS(supplier_description, r'\\d{2}X\\d+(\\.\\d+)?\\s+\\d+X\\d+(\\.\\d+)?\\s+ET-?\\d+')\n" +
  "    OR REGEXP_CONTAINS(UPPER(supplier_description), r'\\bCB\\d+(\\.\\d+)?\\b')";

function stockProductTypesSQL(names) {
  return 'SELECT\n' +
    '  CASE WHEN ' + WHEEL_PATTERN_SQL + ' THEN \'wheel\' ELSE \'tyre_or_other\' END AS product_type,\n' +
    '  match_status,\n' +
    '  COUNT(*) AS n\n' +
    'FROM `' + PROJECT_ID + '.' + STOCK_STOCK_TABLE + '`\n' +
    "WHERE qty > 0 AND supplier_pid NOT LIKE r'DELETED\\_%'\n" +
    '  AND supplier_name IN (' + sqlQuoteList(names) + ')\n' +
    'GROUP BY product_type, match_status';
}

async function fetchProductTypeBreakdown(keys) {
  const names = stockRowsForKeys(keys);
  if (!names.length) return { wheel: { matched: 0, unmatched: 0, total: 0 }, tyre_or_other: { matched: 0, unmatched: 0, total: 0 } };
  const rows = await runBQQuery(stockProductTypesSQL(names));
  const out = { wheel: { matched: 0, unmatched: 0, total: 0 }, tyre_or_other: { matched: 0, unmatched: 0, total: 0 } };
  for (const r of rows) {
    const type = r.product_type === 'wheel' ? 'wheel' : 'tyre_or_other';
    const n = Number(r.n) || 0;
    if (r.match_status === 'matched') out[type].matched += n; else out[type].unmatched += n;
    out[type].total += n;
  }
  return out;
}

function stockMappingSQL() {
  const inList = STOCK_ALLOWED_NAMES.map(n => "'" + n.replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'").join(',');
  return 'SELECT supplier_name, match_status, COUNT(*) AS n\n' +
    'FROM `' + PROJECT_ID + '.' + STOCK_STOCK_TABLE + '`\n' +
    "WHERE qty > 0 AND supplier_pid NOT LIKE r'DELETED\\_%'\n" +
    '  AND supplier_name IN (' + inList + ')\n' +
    'GROUP BY supplier_name, match_status';
}

// Aggregates supplier_stock rows into the {key, matched, unmatched, total, locations[]}
// shape the dashboard expects. Grouping (location -> supplier key) is done here in JS.
async function fetchStockMapping() {
  const rows = await runBQQuery(stockMappingSQL());
  const byKey = {};
  for (const r of rows) {
    const key = STOCK_MAP_NAME_TO_KEY[r.supplier_name];
    if (!key) continue; // outside the curated set
    const n = Number(r.n) || 0;
    const matched = r.match_status === 'matched';
    const g = byKey[key] || (byKey[key] = { key, matched: 0, unmatched: 0, total: 0, locs: {} });
    const loc = g.locs[r.supplier_name] || (g.locs[r.supplier_name] = { name: r.supplier_name, matched: 0, unmatched: 0, total: 0 });
    if (matched) { g.matched += n; loc.matched += n; } else { g.unmatched += n; loc.unmatched += n; }
    g.total += n; loc.total += n;
  }
  return Object.keys(byKey).map(k => {
    const g = byKey[k];
    return {
      key: g.key, matched: g.matched, unmatched: g.unmatched, total: g.total,
      locations: Object.keys(g.locs).map(nm => g.locs[nm]).sort((a, b) => b.total - a.total)
    };
  }).sort((a, b) => b.total - a.total);
}

async function runBQQuery(sql) {
  return await runBQQueryREST(sql);
}

async function runBQQueryCLI(sql) {
  const gc = getGcloud();
  if (!gc) throw new Error('gcloud not found — run Setup_BigQuery_Auth.bat');

  const bqCmd = gc.replace(/gcloud(\.cmd)?$/i, 'bq$1');
  console.log('[BQ-CLI] Using bq:', bqCmd);

  // Write SQL to temp file to avoid shell escaping issues
  const os = require('os');
  const tmpSQL = os.tmpdir().replace(/\\/g,'/') + '/stockmatch_query.sql';
  fs.writeFileSync(tmpSQL, sql, 'utf8');

  console.log('[BQ-CLI] Running query via bq CLI (max_rows=200000)...');
  const jsonOut = await new Promise((res,rej) => {
    // --max_rows=200000 prevents the default 100-row limit
    const cmd = `"${bqCmd}" query --project_id=${PROJECT_ID} --format=json --use_legacy_sql=false --max_rows=200000 < "${tmpSQL}"`;
    cp.exec(cmd, {timeout:300000, maxBuffer:200*1024*1024, windowsHide:true},
      (err,stdout,stderr) => {
        if (err) rej(new Error('bq CLI: '+(stderr||err.message).slice(0,400)));
        else res(stdout);
      });
  });

  let rows;
  try { rows = JSON.parse(jsonOut); }
  catch(e) { throw new Error('bq CLI parse error: '+jsonOut.slice(0,300)); }

  if (!Array.isArray(rows)) throw new Error('bq CLI returned non-array: '+jsonOut.slice(0,200));
  console.log('[BQ-CLI] Got', rows.length, 'rows');
  if (rows.length === 0) throw new Error('bq CLI returned 0 rows — check auth with: gcloud auth application-default print-access-token');

  // bq CLI returns flat objects with column names as keys
  // Normalise field names to match REST API (e.g. 'key' → 'supplier_key' if needed)
  const sample = rows[0] || {};
  console.log('[BQ-CLI] Fields:', Object.keys(sample).join(', '));

  // Map bq CLI field names to our expected names if they differ
  // The SQL uses AS aliases so names should match, but normalise just in case
  return rows.map(r => ({
    supplier_pid:      String(r.supplier_pid      || '').trim(),
    spinach_id:        String(r.spinach_id        || r.spinach_pid || '').trim(),
    data_provider_pid: String(r.data_provider_pid || '').trim(),
    manufacturer:      String(r.manufacturer      || '').trim(),
    profile:           String(r.profile           || '').trim(),
    profile_text:      String(r.profile_text      || '').trim(),
    dimensions:        String(r.dimensions        || '').trim(),
    supplier_name:     String(r.supplier_name     || r.name_supplier || '').trim(),
    supplier_key:      String(r.supplier_key      || r['key'] || '').trim().toLowerCase(),
  }));
}

async function runBQQueryREST(sql) {
  const token = await getADCToken();
  const auth  = {'Authorization':'Bearer '+token, 'Content-Type':'application/json'};
  const BASE  = `/bigquery/v2/projects/${PROJECT_ID}`;

  // ── Submit job ──────────────────────────────────────────────────────────
  console.log('[BQ] Submitting job...');
  const job = await post('bigquery.googleapis.com', `${BASE}/jobs`,
    JSON.stringify({configuration:{query:{query:sql,useLegacySql:false}}}), auth);

  if (!job || job.error) {
    const code = job?.error?.code || 0;
    const msg  = job?.error?.message || 'No response';
    console.error('[BQ] Submit error:', code, msg);
    if (code===401||code===403||code===404||msg.includes('UNAUTHENTICATED')) throw new Error('NOT_AUTHENTICATED');
    throw new Error(`BQ error ${code}: ${msg}`);
  }

  const jobId    = job.jobReference?.jobId;
  const location = job.jobReference?.location || 'US';
  if (!jobId) throw new Error('BQ returned no jobId');
  const jobIdEnc = encodeURIComponent(jobId);
  console.log(`[BQ] Job ${jobId} location=${location}`);

  // ── Poll until done ─────────────────────────────────────────────────────
  for (let i=0; i<80; i++) {
    await new Promise(r=>setTimeout(r, i<5 ? 1500 : 3000));
    const s = await get('bigquery.googleapis.com', `${BASE}/jobs/${jobIdEnc}?location=${location}`, auth);
    if (!s || s.error) throw new Error('BQ poll error: '+(s?.error?.message||'no response'));
    if (i%5===0) console.log(`[BQ] state=${s.status?.state} t=${Math.round((i+1)*2)}s`);
    if (s.status?.state === 'DONE') {
      if (s.status.errorResult) throw new Error('BQ query failed: '+s.status.errorResult.message);
      console.log(`[BQ] Job done after ${i+1} polls`);
      break;
    }
    if (i===79) throw new Error('BQ timeout (2 min)');
  }

  // ── Fetch results ──────────────────────────────────────────────────────
  // Use jobs.getQueryResults (same endpoint but accessed differently)
  // Some proxies block /queryResults sub-path — use the query method instead
  let allRows=[], fields=null, pageToken=null, pageNum=0;

  const parseRows = raws => (raws||[]).map(row=>{
    const o={};row.f.forEach((c,i)=>o[fields[i]]=c.v);return o;
  });

  do {
    pageNum++;
    // Build URL - try both path formats for compatibility
    let url = `${BASE}/queries/${jobIdEnc}?maxResults=10000&timeoutMs=60000&location=${location}`;
    if (pageToken) url += `&pageToken=${encodeURIComponent(pageToken)}`;
    console.log(`[BQ] Fetch page ${pageNum} via /queries/${jobIdEnc.slice(0,20)}...`);

    let pg = await get('bigquery.googleapis.com', url, auth);

    // If /queries/ path fails, fall back to /jobs/{id}/queryResults
    if (!pg || pg.error) {
      console.log(`[BQ] /queries/ path failed (${pg?.error?.code}), trying /jobs/{id}/queryResults...`);
      url = `${BASE}/jobs/${jobIdEnc}/queryResults?maxResults=10000&timeoutMs=60000`;
      if (pageToken) url += `&pageToken=${encodeURIComponent(pageToken)}`;
      pg = await get('bigquery.googleapis.com', url, auth);
    }

    if (!pg || pg.error) throw new Error(`BQ results error: ${pg?.error?.message||'no response'}`);

    if (!fields) {
      if (!pg.schema) throw new Error('BQ returned no schema');
      fields = pg.schema.fields.map(f=>f.name);
      console.log(`[BQ] Total rows: ${pg.totalRows} | fields: ${fields.length}`);
    }

    if (pg.jobComplete === false) {
      console.log(`[BQ] Page ${pageNum}: not ready yet, retrying...`);
      await new Promise(r=>setTimeout(r,3000));
      pageNum--;
      continue;
    }

    const pageRows = parseRows(pg.rows);
    allRows = allRows.concat(pageRows);
    pageToken = pg.pageToken || null;
    console.log(`[BQ] Page ${pageNum}: ${pageRows.length} rows | total: ${allRows.length} | more: ${!!pageToken}`);

  } while (pageToken);

  console.log(`[BQ] Complete: ${allRows.length.toLocaleString()} rows in ${pageNum} pages`);
  if (allRows.length === 0) throw new Error('BQ returned 0 rows');
  return allRows;
}
const snapshotPath = market => path.join(__dirname, `sc_${market}_snapshot.ndjson`);
const RULES_FILE   = path.join(__dirname, 'stockmatch_learned_rules.json');

function saveSnapshotNDJSON(market, rows) {
  const file = snapshotPath(market);
  const lines = rows.map(r=>JSON.stringify({
    supplier_pid:      r.supplier_pid      || '',
    spinach_id:        r.spinach_id        || '',
    data_provider_pid: r.data_provider_pid || '',
    manufacturer:      r.manufacturer      || '',
    profile:           r.profile           || '',
    profile_text:      r.profile_text      || '',
    dimensions:        r.dimensions        || '',
    supplier_name:     r.supplier_name     || '',
    supplier_key:      r.supplier_key      || '',
  })).join('\n');
  fs.writeFileSync(file, lines, 'utf8');
  const sizeMB = (fs.statSync(file).size/1024/1024).toFixed(1);
  return {file, count:rows.length, sizeMB};
}

// ── Helpers ────────────────────────────────────────────────────────────────
function readBody(req) {
  return new Promise((res,rej)=>{
    const chunks=[];
    req.on('data',c=>{chunks.push(c);if(Buffer.concat(chunks).length>50e6)rej(new Error('Too large'));});
    req.on('end',()=>{try{res(JSON.parse(Buffer.concat(chunks).toString()));}catch(e){rej(new Error('Bad JSON'));}});
    req.on('error',rej);
  });
}
function sendJSON(res, code, obj) {
  res.writeHead(code, {'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
  res.end(JSON.stringify(obj));
}

// ── HTTP Server ────────────────────────────────────────────────────────────
const server = http.createServer(async(req,res)=>{
  const url = new URL(req.url, `http://localhost:${PORT}`);
  res.setHeader('Access-Control-Allow-Origin','*');
  res.setHeader('Access-Control-Allow-Methods','GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers','Content-Type');
  if (req.method==='OPTIONS'){res.writeHead(204);res.end();return;}

  // /api/list-paths — list G: drive shortcut targets to find correct paths
  if (req.method==='GET' && url.pathname==='/api/list-paths') {
    const result = {};
    const tryList = (p) => {
      try {
        return fs.readdirSync(p).slice(0,20);
      } catch(e) { return 'ERROR: '+e.message; }
    };
    const tryExists = (p) => {
      try { return fs.existsSync(p); } catch(e) { return false; }
    };

    result['G: exists'] = tryExists('G:\\');
    result['G: contents'] = tryList('G:\\');
    result['.shortcut-targets-by-id exists'] = tryExists('G:\\.shortcut-targets-by-id');
    result['.shortcut-targets-by-id contents'] = tryList('G:\\.shortcut-targets-by-id');

    // Check known shortcut IDs
    const ids = [
      '0BzX9cos2jGfDZzRQeVdsbk9yaXM',  // FTP files (parent of ftp.shipping)
      '0B5Fm-yruEh0AM3FSanoxd3ZnbU0',  // ftp.shipping
      '0B5Fm-yruEh0AdFVZRzFwOTRxWXc',  // Product (parent of Stock_Matcher) - known working
    ];
    for (const id of ids) {
      const p = `G:\\.shortcut-targets-by-id\\${id}`;
      result[`shortcut ${id}`] = tryExists(p) ? tryList(p) : 'NOT FOUND';
    }

    // Try the exact paths we need
    const paths = [
      `G:\\.shortcut-targets-by-id\\0BzX9cos2jGfDZzRQeVdsbk9yaXM\\FTP files\\ftp.shipping\\tyroola_products\\product_data.csv`,
      `G:\\.shortcut-targets-by-id\\0B5Fm-yruEh0AM3FSanoxd3ZnbU0\\tyroola_products\\product_data.xlsx`,
      `G:\\.shortcut-targets-by-id\\0B5Fm-yruEh0AdFVZRzFwOTRxWXc\\Stock_Matcher\\product_data.xlsx`,
    ];
    for (const p of paths) {
      result[p] = tryExists(p) ? 'EXISTS ✓' : 'MISSING';
    }

    sendJSON(res, 200, result);
    return;
  }

  // /api/check-paths — check which Drive paths exist on this machine
  if (req.method==='GET' && url.pathname==='/api/check-paths') {
    const toCheck = {
      'SC AU':  `G:\\.shortcut-targets-by-id\\0B5Fm-yruEh0AdFVZRzFwOTRxWXc\\Stock_Matcher\\RESTORED suppliers_check_big_query - new_file_with_table.xlsx`,
      'SC NZ':  `G:\\.shortcut-targets-by-id\\0B5Fm-yruEh0AdFVZRzFwOTRxWXc\\Stock_Matcher\\NZ Suppliers Check Report BQ.xlsx`,
      'ProdData': `G:\\.shortcut-targets-by-id\\0BzX9cos2jGfDZzRQeVdsbk9yaXM\\FTP files\\ftp.shipping\\tyroola_products\\product_data.csv`,
      'G: drive': `G:\\`,
      'G: shortcut-targets': `G:\\.shortcut-targets-by-id`,
      'Snapshot AU': snapshotPath('au'),
      'Snapshot NZ': snapshotPath('nz'),
      'xlsx pkg': path.join(__dirname,'node_modules','xlsx'),
    };
    const results = {};
    for (const [k,p] of Object.entries(toCheck)) {
      try { results[k] = fs.existsSync(p) ? `EXISTS: ${p}` : `MISSING: ${p}`; }
      catch(e) { results[k] = `ERROR: ${e.message}`; }
    }
    results['__dirname'] = __dirname;
    results['gcloud'] = getGcloud() || 'not found';
    sendJSON(res, 200, results);
    return;
  }

  // /api/version
  if (req.method==='GET' && url.pathname==='/api/version') {
    sendJSON(res,200,{version:'5.1',built:'2026-06-23',project:PROJECT_ID,dir:__dirname});
    return;
  }

  // /api/bq-diag — full diagnostic: token identity + raw BQ test call
  if (req.method==='GET' && url.pathname==='/api/bq-diag') {
    const out = { steps:[] };
    const step = (k,v) => { out.steps.push({k,v}); console.log('[DIAG]',k,':',JSON.stringify(v).slice(0,200)); };

    // 1. Which ADC file exists?
    const adcFile = [
      path.join(process.env.APPDATA||'','gcloud','application_default_credentials.json'),
      path.join(process.env.USERPROFILE||'','AppData','Roaming','gcloud','application_default_credentials.json'),
    ].find(p=>{ try{return fs.existsSync(p);}catch(e){return false;} });
    step('ADC file', adcFile||'NOT FOUND');
    if (adcFile) {
      const adc = JSON.parse(fs.readFileSync(adcFile,'utf8'));
      step('ADC type', adc.type);
      step('ADC client_id', (adc.client_id||'').slice(0,30)+'...');
    }

    // 2. Get token
    let token;
    try {
      token = await getADCToken();
      step('token obtained', token.slice(0,20)+'...');
    } catch(e) {
      step('token ERROR', e.message);
      return sendJSON(res,200,out);
    }

    // 3. Check token identity
    try {
      const info = await get('oauth2.googleapis.com', `/tokeninfo?access_token=${encodeURIComponent(token)}`);
      step('token email', info.email||'unknown');
      step('token scope', (info.scope||'').split(' ').filter(s=>s.includes('bigquery')||s.includes('cloud')));
      step('token expires_in', info.expires_in);
      step('has bigquery scope', (info.scope||'').includes('bigquery')||info.scope==='email');
    } catch(e) { step('tokeninfo ERROR', e.message); }

    // 4. Raw BQ jobs POST
    try {
      const auth = {'Authorization':'Bearer '+token,'Content-Type':'application/json'};
      const testSQL = `SELECT 1 as test`;
      const rawResp = await new Promise((resolve,reject)=>{
        const buf = Buffer.from(JSON.stringify({configuration:{query:{query:testSQL,useLegacySql:false}}}));
        const req2 = require('https').request({
          hostname:'bigquery.googleapis.com',
          path:`/bigquery/v2/projects/${PROJECT_ID}/jobs`,
          method:'POST',
          headers:{'Content-Length':buf.length,...auth}
        }, res2=>{
          const chunks=[];
          res2.on('data',c=>chunks.push(c));
          res2.on('end',()=>resolve({status:res2.statusCode, body:Buffer.concat(chunks).toString('utf8').slice(0,500)}));
        });
        req2.on('error',reject); req2.write(buf); req2.end();
      });
      step('BQ POST status', rawResp.status);
      step('BQ POST response', rawResp.body);
    } catch(e) { step('BQ POST ERROR', e.message); }

    sendJSON(res,200,out);
    return;
  }

  // /api/whoami — token identity check
  if (req.method==='GET' && url.pathname==='/api/whoami') {
    try {
      const t=await getADCToken();
      // Check which account this token belongs to
      const info=await get('oauth2.googleapis.com',`/tokeninfo?access_token=${encodeURIComponent(t)}`);
      sendJSON(res,200,{ok:true,email:info.email,scope:info.scope,token:t.slice(0,15)+'...'});
    } catch(e){sendJSON(res,200,{ok:false,error:e.message});}
    return;
  }

  // /api/health
  if (req.method==='GET' && url.pathname==='/api/health') {
    let authed=false, method='none';
    try{await getADCToken();authed=true;method=getGcloud()?'gcloud':'adc';}catch(e){}
    sendJSON(res,200,{ok:true,project:PROJECT_ID,authenticated:authed,authMethod:method,
      gcloudPath:getGcloud()||'not found'});
    return;
  }

  // /api/log — browser error logging
  if (req.method==='POST' && url.pathname==='/api/log') {
    try{const b=await readBody(req);console.error('[BROWSER]',b.msg,'@',b.src+':'+b.line);}catch(e){}
    sendJSON(res,200,{ok:true});return;
  }

  // /api/test — step by step BQ test
  if (req.method==='GET' && url.pathname==='/api/test') {
    const steps=[];
    const step=(n,ok,d)=>{steps.push({name:n,ok,detail:d});console.log(`[TEST] ${ok?'✓':'✗'} ${n}: ${d}`);};
    const gc=getGcloud();
    step('gcloud found', !!gc, gc||'not found');
    try{
      let token;
      try{token=await getADCToken();step('Auth',true,token.slice(0,20)+'...');}
      catch(e){step('Auth',false,e.message);return sendJSON(res,200,{steps});}
      try{
        const info=await get('oauth2.googleapis.com',`/tokeninfo?access_token=${encodeURIComponent(token)}`);
        step('Token identity',true,`email=${info.email} scope_ok=${(info.scope||'').includes('bigquery')}`);
      }catch(e){step('Token identity',false,e.message);}
      const auth={'Authorization':'Bearer '+token,'Content-Type':'application/json'};
      const sql=`SELECT COUNT(*) as cnt FROM \`${PROJECT_ID}.${DATASET}.${TABLE}\` WHERE stock_level > 0 LIMIT 1`;
      let job;
      try{
        job=await post('bigquery.googleapis.com',`/bigquery/v2/projects/${PROJECT_ID}/jobs`,
          JSON.stringify({configuration:{query:{query:sql,useLegacySql:false}}}),auth);
        if(job.error)throw new Error(`${job.error.code}: ${job.error.message}`);
        step('BQ job submit',true,`jobId=${job.jobReference?.jobId} location=${job.jobReference?.location}`);
      }catch(e){step('BQ job submit',false,e.message);return sendJSON(res,200,{steps});}
      let done2=false;
      for(let i=0;i<20&&!done2;i++){
        await new Promise(r=>setTimeout(r,1500));
        const s=await get('bigquery.googleapis.com',`/bigquery/v2/projects/${PROJECT_ID}/jobs/${job.jobReference.jobId}?location=${job.jobReference.location||'US'}`,auth);
        if(s.status?.state==='DONE'){
          if(s.status.errorResult){step('BQ poll',false,s.status.errorResult.message);return sendJSON(res,200,{steps});}
          step('BQ poll',true,`done in ~${(i+1)*1.5}s`);done2=true;
        }
      }
      if(!done2){step('BQ poll',false,'timeout');return sendJSON(res,200,{steps});}
      const r2=await get('bigquery.googleapis.com',
        `/bigquery/v2/projects/${PROJECT_ID}/jobs/${job.jobReference.jobId}/queryResults?maxResults=1`,auth);
      step('BQ results',!r2.error,r2.error?r2.error.message:`count=${r2.rows?.[0]?.f?.[0]?.v}`);
    }catch(e){steps.push({name:'exception',ok:false,detail:e.message});}
    sendJSON(res,200,{steps});return;
  }

  // /api/prod — read product_data.csv from Drive G: server-side
  if (req.method==='GET' && url.pathname==='/api/prod') {
    // Confirmed exact path from user:
    // G:\.shortcut-targets-by-id\0BzX9cos2jGfDZzRQeVdsbk9yaXM\FTP files\ftp.shipping\tyroola_products\product_data.csv
    const PROD_BASE = `G:\\.shortcut-targets-by-id\\0BzX9cos2jGfDZzRQeVdsbk9yaXM\\FTP files\\ftp.shipping\\tyroola_products`;
    const prodPaths = [
      `${PROD_BASE}\\product_data.csv`,
      `${PROD_BASE}\\product_data.xlsx`,
    ];

    console.log('[PROD] Searching for product_data in:', PROD_BASE);
    prodPaths.forEach(p => {
      try { console.log(' ', fs.existsSync(p) ? 'FOUND:' : 'missing:', p); }
      catch(e) {}
    });

        const localPath = prodPaths.find(p=>{ try{return fs.existsSync(p);}catch(e){return false;} });
    if (!localPath) {
      return sendJSON(res,404,{ok:false,
        error:'product_data.xlsx not found on Drive G:\nChecked:\n'+prodPaths.join('\n')+
        '\n\nOpen http://localhost:'+PORT+'/api/list-paths to see your G: drive structure'});
    }
    try {
      const sizeMB=(fs.statSync(localPath).size/1024/1024).toFixed(1);
      const isCsv=localPath.toLowerCase().endsWith('.csv');
      console.log(`[PROD] Reading ${sizeMB}MB ${isCsv?'CSV':'XLSX'}: ${localPath}`);
      const mtime=fs.statSync(localPath).mtime.toISOString();

      let data=[];

      if(isCsv){
        // Parse CSV line by line
        const content=fs.readFileSync(localPath,'utf8');
        const lines=content.split(/\r?\n/);
        if(!lines.length) throw new Error('Empty CSV file');

        // Parse CSV with quote handling
        const parseCSVLine=line=>{
          const res=[];let cur='';let inQ=false;
          for(let i=0;i<line.length;i++){
            const c=line[i];
            if(c==='"'&&!inQ){inQ=true;}
            else if(c==='"'&&inQ&&line[i+1]==='"'){cur+='"';i++;}
            else if(c==='"'&&inQ){inQ=false;}
            else if(c===','&&!inQ){res.push(cur);cur='';}
            else{cur+=c;}
          }
          res.push(cur);
          return res;
        };

        // Find header row (contains dp_id)
        let hdrRow=-1;
        for(let i=0;i<Math.min(lines.length,10);i++){
          const row=parseCSVLine(lines[i]);
          if(row.some(v=>v.trim().toLowerCase()==='dp_id')){hdrRow=i;break;}
        }
        if(hdrRow===-1) hdrRow=0;

        const hdrs=parseCSVLine(lines[hdrRow]).map(h=>h.trim().toLowerCase().replace(/[^a-z0-9_]/g,'_'));
        data=lines.slice(hdrRow+1).filter(l=>l.trim()).map(l=>{
          const row=parseCSVLine(l);
          const o={};hdrs.forEach((h,i)=>o[h]=String(row[i]??'').trim());
          return o;
        }).filter(r=>r.dp_id||r.brand);

        // Return only raw CSV — browser passes it to loadProd() via Papa.parse
        // Do NOT include parsed rows too — that would double the response size (68MB+)
        console.log(`[PROD] ${data.length} rows, sending ${(content.length/1024/1024).toFixed(1)}MB CSV`);
        sendJSON(res,200,{ok:true,count:data.length,sizeMB,mtime,path:localPath,
          csv:content});
        return;
      } else {
        // Parse XLSX
        const xlsx2=getXLSX();
        if(!xlsx2) return sendJSON(res,500,{ok:false,error:'xlsx library not available. Run Install_StockMatch.bat to install it.'});
        const buf=fs.readFileSync(localPath);
        const wb=xlsx2.read(buf,{type:'buffer'});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw0=xlsx2.utils.sheet_to_json(ws,{header:1,defval:''});
        let hdrRow=0;
        for(let i=0;i<Math.min(raw0.length,10);i++){
          if(raw0[i].some(v=>String(v).toLowerCase().trim()==='dp_id')){hdrRow=i;break;}
        }
        const rows=xlsx2.utils.sheet_to_json(ws,{header:1,range:hdrRow,defval:'',raw:false});
        const hdrs=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[^a-z0-9_]/g,'_'));
        data=rows.slice(1).map(r=>{const o={};hdrs.forEach((h,i)=>o[h]=String(r[i]??'').trim());return o;})
          .filter(r=>r.dp_id||r.brand);
      }

      if(!data.length) throw new Error('No rows with dp_id found in file');
      console.log(`[PROD] ${data.length} rows loaded`);

      // Convert to CSV so frontend uses the same loadProd() → Papa.parse() path
      const allHdrs=Object.keys(data[0]);
      const csvLines=[allHdrs.join(','),...data.map(r=>allHdrs.map(h=>JSON.stringify(String(r[h]||''))).join(','))];
      const csvContent=csvLines.join('\n');
      sendJSON(res,200,{ok:true,count:data.length,sizeMB,mtime,path:localPath,csv:csvContent});
    }catch(e){
      console.error('[PROD] Error:',e.message);
      sendJSON(res,500,{ok:false,error:'Product data read failed: '+e.message});
    }
    return;
  }

  // /api/sc-xlsx — read XLSX from Drive G: server-side, return rows as JSON
  if (req.method==='GET' && url.pathname==='/api/sc-xlsx') {
    const market = (url.searchParams.get('market')||'au').toLowerCase();
    const localPaths = {
      au: `G:\\.shortcut-targets-by-id\\0B5Fm-yruEh0AdFVZRzFwOTRxWXc\\Stock_Matcher\\RESTORED suppliers_check_big_query - new_file_with_table.xlsx`,
      nz: `G:\\.shortcut-targets-by-id\\0B5Fm-yruEh0AdFVZRzFwOTRxWXc\\Stock_Matcher\\NZ Suppliers Check Report BQ.xlsx`,
    };
    const localPath = localPaths[market];
    if (!localPath || !fs.existsSync(localPath)) {
      return sendJSON(res, 404, {ok:false, error:'Drive file not found: '+(localPath||'unknown market')});
    }
    try {
      const sizeMB = (fs.statSync(localPath).size/1024/1024).toFixed(1);
      console.log(`[XLSX] Reading Drive file ${sizeMB}MB: ${localPath}`);

      // Load xlsx — try multiple locations
      let xlsx;
      const xlsxPaths = [
        path.join(__dirname, 'node_modules', 'xlsx', 'lib', 'xlsx.js'),
        path.join(__dirname, 'node_modules', 'xlsx', 'xlsx.js'),
        path.join(__dirname, 'lib', 'xlsx.full.min.js'),
      ];
      for (const p of xlsxPaths) {
        if (fs.existsSync(p)) { try { xlsx = require(p); break; } catch(e) {} }
      }
      if (!xlsx) {
        try { xlsx = require('xlsx'); } catch(e) {
          // xlsx not available — install it
          console.log('[XLSX] Installing xlsx package...');
          cp.execSync('npm install xlsx --prefix "'+__dirname+'"', {timeout:60000, stdio:'pipe'});
          xlsx = require(path.join(__dirname, 'node_modules', 'xlsx', 'xlsx.js'));
        }
      }
      const buf = fs.readFileSync(localPath);
      const wb  = xlsx.read(buf, {type:'buffer'});
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = xlsx.utils.sheet_to_json(ws, {defval:'', raw:false});
      const rows = raw.map(r=>({
        supplier_pid:      String(r['Supplier PID']      ||r.supplier_pid      ||'').trim(),
        spinach_id:        String(r['Spinach ID']        ||r.spinach_id        ||r.spinach_pid||'').trim(),
        data_provider_pid: String(r['Data Provider PID'] ||r.data_provider_pid ||'').trim(),
        manufacturer:      String(r['Manufacturer']      ||r.manufacturer      ||'').trim(),
        profile:           String(r['Profile']           ||r.profile           ||'').trim(),
        profile_text:      String(r['Profile Text']      ||r.profile_text      ||'').trim(),
        dimensions:        String(r['Dimensions']        ||r.dimensions        ||'').trim(),
        supplier_name:     String(r['Supplier Name']     ||r.supplier_name     ||r.name_supplier||'').trim(),
        supplier_key:      String(r['Supplier Key']      ||r.supplier_key      ||r['key']||'').trim().toLowerCase(),
      })).filter(r=>r.supplier_pid||r.manufacturer);
      console.log(`[XLSX] ${rows.length} rows from Drive file`);
      // Also save as NDJSON snapshot for fast subsequent loads
      saveSnapshotNDJSON(market, rows);
      const mtime=fs.statSync(localPath).mtime.toISOString();
      sendJSON(res, 200, {ok:true, rows, count:rows.length, sizeMB, mtime});
    } catch(e) {
      sendJSON(res, 500, {ok:false, error:'Drive XLSX read failed: '+e.message});
    }
    return;
  }

  // /api/sc-test — run full BQ query and return detailed step results
  if (req.method==='GET' && url.pathname==='/api/sc-test') {
    const steps=[];
    const log=(k,v)=>{steps.push({k,v});console.log('[SC-TEST]',k,':',JSON.stringify(v).slice(0,200));};
    try{
      const token=await getADCToken();
      log('token',token.slice(0,20)+'...');
      const auth={'Authorization':'Bearer '+token,'Content-Type':'application/json'};
      const BASE=`/bigquery/v2/projects/${PROJECT_ID}`;

      // Submit
      const sql=`SELECT COUNT(*) as cnt FROM \`${PROJECT_ID}.${DATASET}.${TABLE}\` WHERE is_active=true AND is_active_stock=true LIMIT 1`;
      log('sql',sql.slice(0,80));
      const job=await post('bigquery.googleapis.com',`${BASE}/jobs`,
        JSON.stringify({configuration:{query:{query:sql,useLegacySql:false}}}),auth);
      if(job.error){log('submit_error',job.error);return sendJSON(res,200,{ok:false,steps});}
      const jobId=job.jobReference?.jobId;
      const location=job.jobReference?.location||'US';
      log('job_submitted',{jobId,location});

      // Poll
      let done=false;
      for(let i=0;i<20&&!done;i++){
        await new Promise(r=>setTimeout(r,2000));
        const pollUrl=`${BASE}/jobs/${jobId}?location=${location}`;
        log('polling',pollUrl.slice(0,80));
        const s=await get('bigquery.googleapis.com',pollUrl,auth);
        if(s.error){log('poll_error',s.error);return sendJSON(res,200,{ok:false,steps});}
        log('poll_state',s.status?.state);
        if(s.status?.state==='DONE'){
          if(s.status.errorResult){log('job_failed',s.status.errorResult);return sendJSON(res,200,{ok:false,steps});}
          done=true;
        }
      }
      if(!done){log('timeout','job did not complete');return sendJSON(res,200,{ok:false,steps});}

      // Fetch results
      const resUrl=`${BASE}/jobs/${jobId}/queryResults?maxResults=10`;
      log('fetch_results',resUrl.slice(0,80));
      const r2=await get('bigquery.googleapis.com',resUrl,auth);
      if(r2.error){log('results_error',r2.error);return sendJSON(res,200,{ok:false,steps});}
      log('total_rows',r2.totalRows);
      log('rows',r2.rows);
      log('SUCCESS','BQ working correctly');
      sendJSON(res,200,{ok:true,steps,totalRows:r2.totalRows});
    }catch(e){
      steps.push({k:'exception',v:e.message});
      sendJSON(res,200,{ok:false,steps,error:e.message});
    }
    return;
  }

  // /api/sc-reset — delete stale snapshots so fresh BQ data is loaded
  if (req.method==='GET' && url.pathname==='/api/sc-reset') {
    const results={};
    ['au','nz'].forEach(m=>{
      const f=snapshotPath(m);
      if(fs.existsSync(f)){
        try{fs.unlinkSync(f);results[m]='deleted';}
        catch(e){results[m]='error: '+e.message;}
      } else {
        results[m]='not found';
      }
    });
    console.log('[SC] Snapshots reset:',results);
    sendJSON(res,200,{ok:true,results,message:'Old snapshots deleted. Refresh AU from BQ to get fresh data.'});
    return;
  }

  // /api/sc — query BQ (AU only), save snapshot, return count only
  // NZ uses Drive XLSX — the NZ BQ table references cross-project tables we can't access
  if (req.method==='GET' && url.pathname==='/api/sc') {
    const market=(url.searchParams.get('market')||'au').toLowerCase();
    const snapFile=snapshotPath(market);

    // NZ: query BQ live using suppliers_check_report_table (materialized replica — no popeye_production access needed)
    if (market === 'nz') {
      console.log('[SC] NZ market: querying BQ live (suppliers_check_report_table)');
      try {
        const t=Date.now();
        const rows=await runBQQuery(NZ_SC_SQL);
        if(!rows||rows.length===0) throw new Error('NZ BQ returned 0 rows — check filter or auth');
        const{file,count,sizeMB}=saveSnapshotNDJSON(market,rows);
        const elapsed=((Date.now()-t)/1000).toFixed(1);
        console.log(`[BQ] NZ snapshot: ${count} rows, ${sizeMB}MB saved (${elapsed}s)`);
        sendJSON(res,200,{ok:true,count,sizeMB,elapsed,market,source:'bq_live'});
      } catch(e) {
        console.error('[BQ] NZ error:',e.message,'— falling back to Drive XLSX');
        // Fallback: Drive XLSX
        const nzPaths = [
          `G:\\.shortcut-targets-by-id\\0B5Fm-yruEh0AdFVZRzFwOTRxWXc\\Stock_Matcher\\NZ Suppliers Check Report BQ.xlsx`,
        ];
        const nzPath = nzPaths.find(p=>{ try{return fs.existsSync(p);}catch(e2){return false;} });
        if (!nzPath) {
          return sendJSON(res,500,{ok:false,error:'NZ BQ failed and Drive XLSX fallback not found. BQ error: '+e.message});
        }
        try {
          const sizeMBx=(fs.statSync(nzPath).size/1024/1024).toFixed(1);
          console.log(`[XLSX] NZ fallback ${sizeMBx}MB: ${nzPath}`);
          const xlsx=getXLSX();
          if(!xlsx) return sendJSON(res,500,{ok:false,error:'xlsx library not available. Run Install_StockMatch.bat.'});
          const buf=fs.readFileSync(nzPath);
          const wb=xlsx.read(buf,{type:'buffer'});
          const ws=wb.Sheets[wb.SheetNames[0]];
          const arrRows=xlsx.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
          if(!arrRows.length) return sendJSON(res,404,{ok:false,error:'NZ XLSX appears empty'});
          const headers=arrRows[0].map((h,i)=>{
            const s=String(h).trim();
            if(i===0){ const m2=s.match(/\s(\S+)$/); return m2?m2[1]:s; }
            return s;
          });
          console.log(`[XLSX] NZ headers (first 10): ${headers.slice(0,10).join(', ')}`);
          const filteredRows=arrRows.slice(1).reduce((acc,arr)=>{
            const r={};
            headers.forEach((h,i)=>{ r[h]=arr[i]!==undefined?String(arr[i]):''; });
            const active=(r.is_active||'').toLowerCase();
            const activeStock=(r.is_active_stock||'').toLowerCase();
            if(active==='true' && activeStock==='true'){
              const mapped={
                supplier_pid:      (r.supplier_pid||'').trim(),
                spinach_id:        (r.spinach_pid||r.spinach_id||'').trim(),
                data_provider_pid: (r.data_provider_pid||'').trim(),
                manufacturer:      (r.manufacturer||'').trim(),
                profile:           (r.profile||'').trim(),
                profile_text:      (r.profile_text||'').trim(),
                dimensions:        (r.dimensions||'').trim(),
                supplier_name:     (r.name_supplier||r.supplier_name||'').trim(),
                supplier_key:      (r['key']||r.supplier_key||'').trim().toLowerCase(),
              };
              if(mapped.supplier_pid||mapped.manufacturer||mapped.spinach_id) acc.push(mapped);
            }
            return acc;
          },[]);
          const{file,count,sizeMB:sMB}=saveSnapshotNDJSON(market,filteredRows);
          const mtime=fs.statSync(nzPath).mtime.toISOString();
          console.log(`[SC] NZ XLSX fallback snapshot: ${count} rows, ${sMB}MB saved`);
          sendJSON(res,200,{ok:true,count,sizeMB:sMB,market,source:'drive_xlsx_fallback',mtime});
        } catch(e2) {
          sendJSON(res,500,{ok:false,error:'NZ BQ failed and XLSX fallback also failed: '+e2.message});
        }
      }
      return;
    }

    // AU: query BQ live
    try{
      console.log(`\n[BQ] === Refresh SC AU ===`);
      const t=Date.now();
      const rows=await runBQQuery(SC_SQL);
      if(!rows||rows.length===0) throw new Error('BQ returned 0 rows — check filter or auth');
      const{file,count,sizeMB}=saveSnapshotNDJSON(market,rows);
      const elapsed=((Date.now()-t)/1000).toFixed(1);
      console.log(`[BQ] Snapshot: ${count} rows, ${sizeMB}MB saved (${elapsed}s)`);
      sendJSON(res,200,{ok:true,count,sizeMB,elapsed,market});
    }catch(e){
      console.error('[BQ] Error:',e.message);
      // Delete stale snapshot so old data doesn't mislead subsequent page loads
      try{ if(fs.existsSync(snapFile)) fs.unlinkSync(snapFile); }catch(e2){}
      const isAuth=e.message==='NOT_AUTHENTICATED'||e.message.includes('UNAUTHENTICATED')||
                   e.message.includes('invalid_rapt')||e.message.includes('reauth')||
                   e.message.includes('NOT_AUTHENTICATED');
      sendJSON(res,isAuth?401:500,{ok:false,
        error:isAuth?'NOT_AUTHENTICATED':e.message,
        hint:isAuth?'Run Setup_BigQuery_Auth.bat then restart StockMatch':undefined});
    }
    return;
  }

  // /api/rules GET — load rules from disk (called on app startup)
  if (req.method==='GET' && url.pathname==='/api/rules') {
    try {
      if (fs.existsSync(RULES_FILE)) {
        const rules=JSON.parse(fs.readFileSync(RULES_FILE,'utf8'));
        return sendJSON(res,200,{ok:true,rules,file:RULES_FILE});
      }
      return sendJSON(res,200,{ok:true,rules:[],file:RULES_FILE});
    } catch(e) {
      return sendJSON(res,500,{ok:false,error:'Failed to load rules: '+e.message});
    }
  }

  // /api/rules POST — save rules to disk (called whenever rules change)
  if (req.method==='POST' && url.pathname==='/api/rules') {
    try {
      const body=await readBody(req);
      const rules=Array.isArray(body.rules)?body.rules:[];
      fs.writeFileSync(RULES_FILE,JSON.stringify(rules,null,2),'utf8');
      console.log(`[Rules] Saved ${rules.length} rule(s) → ${RULES_FILE}`);
      return sendJSON(res,200,{ok:true,count:rules.length,file:RULES_FILE});
    } catch(e) {
      console.error('[Rules] Save error:',e.message);
      return sendJSON(res,500,{ok:false,error:'Failed to save rules: '+e.message});
    }
  }

  // /api/lmm — proxy LMM mismatch-check requests to Anthropic (avoids CORS; key comes from the browser per-request, never stored server-side)
  if (req.method==='POST' && url.pathname==='/api/lmm') {
    try {
      const body=await readBody(req);
      const apiKey=body.apiKey;
      if (!apiKey) {
        return sendJSON(res,400,{ok:false,error:'No API key provided. Enter your Anthropic key in the LMM Check panel.'});
      }
      const antBody=JSON.stringify({
        model:body.model||'claude-3-5-haiku-20241022',
        max_tokens:body.max_tokens||600,
        messages:body.messages||[],
      });
      // 20s server-side timeout — if Anthropic hangs, abort and return error so browser retry logic can fire
      const j=await Promise.race([
        post('api.anthropic.com','/v1/messages',antBody,{
          'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01',
        }),
        new Promise((_,rej)=>setTimeout(()=>rej(new Error('Anthropic request timed out after 20s')),20000)),
      ]);
      if (j.error) {
        return sendJSON(res,502,{ok:false,error:'Anthropic: '+j.error.message});
      }
      const text=(j.content||[]).filter(b=>b.type==='text').map(b=>b.text).join('');
      sendJSON(res,200,{ok:true,text});
    } catch(e) {
      console.error('[LMM] Error:',e.message);
      sendJSON(res,500,{ok:false,error:e.message});
    }
    return;
  }

  // /api/sc-page — stream snapshot page by page via readline (no full file parse)
  if (req.method==='GET' && url.pathname==='/api/sc-page') {
    const market=(url.searchParams.get('market')||'au').toLowerCase();
    const page=parseInt(url.searchParams.get('page')||'0');
    const size=parseInt(url.searchParams.get('size')||'2000');
    const file=snapshotPath(market);

    if (!fs.existsSync(file)) {
      return sendJSON(res,404,{ok:false,error:'Snapshot not found — click Refresh AU from BQ first'});
    }
    try{
      const start=page*size, end=start+size;
      const rows=[]; let total=0;
      await new Promise((resolve,reject)=>{
        const iface=rl_mod.createInterface({
          input:fs.createReadStream(file,{encoding:'utf8'}),crlfDelay:Infinity});
        iface.on('line',line=>{
          if(!line.trim())return;
          if(total>=start&&total<end){try{rows.push(JSON.parse(line));}catch(e){}}
          total++;
        });
        iface.on('close',resolve);
        iface.on('error',reject);
      });
      const stat = fs.statSync(file);
      sendJSON(res,200,{ok:true,page,total,rows,done:start+rows.length>=total,
        mtime:stat.mtime.toISOString(),sizeMB:(stat.size/1024/1024).toFixed(1)});
    }catch(e){
      sendJSON(res,500,{ok:false,error:'Snapshot read error: '+e.message});
    }
    return;
  }

  // /lib/* — serve local JS libraries
  if (req.method==='GET' && url.pathname.startsWith('/lib/')) {
    const file=path.join(__dirname,'lib',path.basename(url.pathname));
    if (fs.existsSync(file)){res.writeHead(200,{'Content-Type':'application/javascript'});fs.createReadStream(file).pipe(res);}
    else{res.writeHead(404);res.end('Library not found. Run Install_StockMatch.bat to download it.');}
    return;
  }

  // /api/stock-mapping — live supplier stock-mapping data for the dashboard
  if (req.method==='GET' && url.pathname==='/api/stock-mapping') {
    try {
      const data = await fetchStockMapping();
      let refreshed_at = null;
      try {
        const t = await runBQQuery('SELECT CAST(MAX(refreshed_at) AS STRING) AS refreshed_at FROM `' + PROJECT_ID + '.' + STOCK_STOCK_TABLE + '`');
        refreshed_at = (t[0] && t[0].refreshed_at) || null;
      } catch(e) { /* timestamp is best-effort */ }
      const totals = data.reduce((a,d)=>{ a.matched+=d.matched; a.unmatched+=d.unmatched; a.total+=d.total; return a; }, {matched:0,unmatched:0,total:0});
      sendJSON(res, 200, { ok:true, generated_at:new Date().toISOString(), refreshed_at, suppliers:data.length, totals, data });
    } catch(e) {
      const isAuth = /NOT_AUTHENTICATED|UNAUTHENTICATED|invalid_grant|reauth/i.test(e.message||'');
      sendJSON(res, isAuth?401:500, { ok:false, error:e.message, hint: isAuth ? 'Run Setup_BigQuery_Auth.bat (gcloud auth application-default login), then restart StockMatch.' : undefined });
    }
    return;
  }

  // /api/stock-mapping/rows — row-level export (CSV/Excel) for selected supplier(s) + status
  if (req.method==='GET' && url.pathname==='/api/stock-mapping/rows') {
    try {
      const keysParam = (url.searchParams.get('keys')||'').trim();
      const statusParam = (url.searchParams.get('status')||'all').trim().toLowerCase();
      const status = (statusParam==='matched'||statusParam==='unmatched') ? statusParam : 'all';
      let keys = keysParam ? keysParam.split(',').map(s=>s.trim()).filter(Boolean) : STOCK_KNOWN_KEYS.slice();
      keys = keys.filter(k => STOCK_KNOWN_KEYS.includes(k));
      if (!keys.length) { sendJSON(res, 400, {ok:false, error:'No valid supplier keys given. Known keys: '+STOCK_KNOWN_KEYS.join(', ')}); return; }
      const rows = await fetchStockMappingRows(keys, status);
      sendJSON(res, 200, { ok:true, generated_at:new Date().toISOString(), filters:{keys, status}, count:rows.length, rows });
    } catch(e) {
      const isAuth = /NOT_AUTHENTICATED|UNAUTHENTICATED|invalid_grant|reauth/i.test(e.message||'');
      sendJSON(res, isAuth?401:500, { ok:false, error:e.message, hint: isAuth ? 'Run Setup_BigQuery_Auth.bat (gcloud auth application-default login), then restart StockMatch.' : undefined });
    }
    return;
  }

  // /api/stock-mapping/product-types — wheel-vs-tyre breakdown for selected supplier(s)
  if (req.method==='GET' && url.pathname==='/api/stock-mapping/product-types') {
    try {
      const keysParam = (url.searchParams.get('keys')||'').trim();
      let keys = keysParam ? keysParam.split(',').map(s=>s.trim()).filter(Boolean) : [];
      keys = keys.filter(k => STOCK_KNOWN_KEYS.includes(k));
      if (!keys.length) { sendJSON(res, 400, {ok:false, error:'No valid supplier keys given. Known keys: '+STOCK_KNOWN_KEYS.join(', ')}); return; }
      const breakdown = await fetchProductTypeBreakdown(keys);
      sendJSON(res, 200, { ok:true, generated_at:new Date().toISOString(), filters:{keys}, breakdown });
    } catch(e) {
      const isAuth = /NOT_AUTHENTICATED|UNAUTHENTICATED|invalid_grant|reauth/i.test(e.message||'');
      sendJSON(res, isAuth?401:500, { ok:false, error:e.message, hint: isAuth ? 'Run Setup_BigQuery_Auth.bat (gcloud auth application-default login), then restart StockMatch.' : undefined });
    }
    return;
  }

  // Live stock-mapping dashboard page (must precede the .html catch-all below)
  if (req.method==='GET' && url.pathname==='/supplier_stock_match_dashboard_live.html') {
    const f = path.join(__dirname, 'supplier_stock_match_dashboard_live.html');
    if (fs.existsSync(f)){res.writeHead(200,{'Content-Type':'text/html; charset=utf-8'});res.end(fs.readFileSync(f));}
    else{res.writeHead(404);res.end('supplier_stock_match_dashboard_live.html not found in: '+__dirname);}
    return;
  }

  // / — serve app HTML
  if (req.method==='GET' && (url.pathname==='/'||url.pathname.endsWith('.html'))) {
    if (fs.existsSync(APP_FILE)){res.writeHead(200,{'Content-Type':'text/html; charset=utf-8'});res.end(fs.readFileSync(APP_FILE));}
    else{res.writeHead(404);res.end('stock_matcher_app.html not found in: '+__dirname);}
    return;
  }

  res.writeHead(404);res.end('Not found');
});

server.on('error',e=>{
  if(e.code==='EADDRINUSE'){const p=Number(PORT)+1;console.log(`Port ${PORT} busy, trying ${p}`);server.listen(p,'127.0.0.1');}
  else throw e;
});

server.listen(PORT,'127.0.0.1',async()=>{
  const p=server.address().port;
  console.log(`\n  ╔══════════════════════════════════════════════╗`);
  console.log(`  ║  StockMatch v5.3  →  http://localhost:${p}    ║`);
  console.log(`  ╚══════════════════════════════════════════════╝`);
  console.log(`  Project : ${PROJECT_ID}`);

  // Delete stale snapshots on startup so old data never persists across restarts
  ['au','nz'].forEach(m=>{
    const f=snapshotPath(m);
    if(fs.existsSync(f)){
      try{
        const stat=fs.statSync(f);
        const lines=fs.readFileSync(f,'utf8').split('\n').filter(l=>l.trim()).length;
        if(lines<1000){
          fs.unlinkSync(f);
          console.log(`  Deleted stale snapshot (${lines} rows): sc_${m}_snapshot.ndjson`);
        } else {
          console.log(`  Snapshot sc_${m}_snapshot.ndjson: ${lines.toLocaleString()} rows (${(stat.size/1024/1024).toFixed(1)}MB)`);
        }
      }catch(e){}
    }
  });

  const gc=getGcloud();
  console.log(`  gcloud  : ${gc||'NOT FOUND — install Google Cloud SDK'}`);

  // Test auth AND BQ access on startup
  try{
    const token = await getADCToken();
    console.log(`  Token   : ✓ obtained (${token.slice(0,20)}...)`);

    // Verify token actually works for BQ (catches scope issues early)
    const auth = {'Authorization':'Bearer '+token,'Content-Type':'application/json'};
    const testJob = await post('bigquery.googleapis.com',
      `/bigquery/v2/projects/${PROJECT_ID}/jobs`,
      JSON.stringify({configuration:{query:{query:'SELECT 1',useLegacySql:false}}}),
      auth);

    if(testJob.error){
      const code = testJob.error.code;
      const msg = testJob.error.message;
      console.log(`  BQ test : ✗ Error ${code}: ${msg}`);
      if(code===404||code===403||code===401){
        console.log(`\n  ┌─────────────────────────────────────────────────────────┐`);
        console.log(`  │  TOKEN LACKS BIGQUERY SCOPE — run this command:         │`);
        console.log(`  │                                                         │`);
        console.log(`  │  gcloud auth application-default login                  │`);
        console.log(`  │   --scopes=https://www.googleapis.com/auth/bigquery,    │`);
        console.log(`  │            https://www.googleapis.com/auth/cloud-platform│`);
        console.log(`  │                                                         │`);
        console.log(`  │  Or run: Setup_BigQuery_Auth.bat                        │`);
        console.log(`  └─────────────────────────────────────────────────────────┘\n`);
      }
    } else {
      console.log(`  BQ test : ✓ BigQuery access confirmed`);
      console.log('\n  Ready — open StockMatch and click Live BQ Refresh AU\n');
    }
  }catch(e){
    console.log(`  Auth    : ✗ ${e.message}\n`);
    console.log(`  Fix     : Run Setup_BigQuery_Auth.bat\n`);
  }
  console.log('  Ctrl+C to stop\n');
});

