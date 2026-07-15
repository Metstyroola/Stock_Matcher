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

// Wheel/tyre/other classification for a supplier's stock descriptions.
// Validated 2026-07-15 against Tempetyres + Chapel Corner by direct sampling of
// matched descriptions (see docs in memory) — treat any OTHER supplier's split
// as unverified/directional until independently spot-checked; several false
// positives were found and fixed during validation (Michelin's "NNN/NN ZRNN"
// space+Z format, empty/numeric-only descriptions defaulted to tyre rather than
// other since this is a tyre-stock feed and missing text isn't evidence of a
// different product).
const WHEEL_PATTERN_SQL =
  "REGEXP_CONTAINS(UPPER(supplier_description), r'\\bWHEEL[S]?\\b|\\bRIM[S]?\\b|\\bALLOY\\b|\\bWHL\\b|\\bHUB\\b|\\bSPOKE[S]?\\b|\\bSTEELIE[S]?\\b|\\bMAG\\b')\n" +
  "    OR REGEXP_CONTAINS(supplier_description, r'\\d{2}X\\d+(\\.\\d+)?\\s+\\d+X\\d+(\\.\\d+)?\\s+ET-?\\d+')\n" +
  "    OR REGEXP_CONTAINS(UPPER(supplier_description), r'\\bCB\\d+(\\.\\d+)?\\b')";

const NO_DESC_PATTERN_SQL =
  "TRIM(IFNULL(supplier_description,'')) = ''\n" +
  "    OR REGEXP_CONTAINS(TRIM(supplier_description), r'^[0-9]+$')";

const TYRE_PATTERN_SQL =
  "REGEXP_CONTAINS(UPPER(supplier_description), r'\\d{2,3}\\s*/\\s*\\d{2,3}\\s*Z?R\\s*\\d{2}')\n" +
  "    OR REGEXP_CONTAINS(UPPER(supplier_description), r'\\d{2}(\\.\\d+)?X\\d{2}(\\.\\d+)?R\\d{2}')\n" +
  "    OR REGEXP_CONTAINS(UPPER(supplier_description), r'\\bTYRE[S]?\\b|\\bTIRE[S]?\\b')\n" +
  "    OR REGEXP_CONTAINS(supplier_description, r'^\\d{7}[A-Z]')";

const CHAPELCORNER_EXCLUSION_LIST = ["PLASTIC FERRULE (EACH)","TR414 - VALVE COMPLETE (EACH)","TR415 - VALVE COMPLETE (EACH)","TR412 - VALVE COMPLETE (EACH)","TR600HP HP VALVES (EACH)","TR618A - TYRE VALVE METAL SCREW-IN each","TR413 - VALVE AC CHROME (each)","GTS POWER BALANCE PB 2 (1/bag)","GTS POWER BALANCE PB 4 (1/bag)","TPMS VALVE NISSAN FITMENT (6-208)","TR618T - TYRE VALVE RUBBER BASE (each)","TPMS VALVE 1520 (EACH)","PLASTIC CBL RING 108 - 100.1mm","METAL VALVE EXTENSION MEX127 (CHROME)","GTS CEMENT TG-C 10ML TUBE","TPMS VALVE NEW FORD FITMENT (6-210H)","PLASTIC VALVE EXTENSION EX115 (BLACK)","STICK-ON Pb (25g*4) 30/BOX","GTS POWER BALANCE PB 1.5 (3)","PLASTIC CBL RING 108 - 67.1mm","PLASTIC CBL RING 108 - 106.1mm","EXTREME CBL RING 106 - 93.2mm","VALVE CORE TRC1R RED (100/BAG)","METAL VALVE EXTENSION MEX77 (BRASS)","SCREW-IN DOME 6/139.7 108mm x 70mm BLACK (WITH EXTREME LOGO)","VERMAR LIGHT CHEMICAL CEMENT 250ML","GTS POWER BALANCE PB 3 (2/bag)","TRUCK VALVE V3.20.4","1 X TG2-05 BIAS PLY 127X127MM each (ALSO FROM BULK PACK)","REPAIR STRING TRUCK LONG 200MM (25)","FAST FLOW ADAPTOR - AG. TYPE","MS-70-7 TRUCK VALVE (EACH)","EXTREME CBL RING 106 - 67.1mm","CHROME DOME SMALL 75mm TRAILER","J-SHAPE METAL VALVE EXTENSION","VALVE CORE TOOL SINGLE","EXTREME 5/150 CENTER CAPS / DOMES BLACK w/ RED WRITING","GTS ATV REPAIR KIT Q647 - 20pcs (CLEAR BOX)","VS-8-45 DEGREE VALVE (EACH)","TR573 - TYRE VALVE (EACH)","FLEXIBLE VALVE EXTENSION 105mm","WHEEL NUT INDICATOR 41MM (BAG OF 25) YELLOW","EXTREME 6/139.7 CENTER CAPS / DOMES BLACK w/ WHITE WRITING","SPLINE LOCK SOCKET - LARGE 7 SPLINE SLOTTED (NEW) K596L","MAG NUT CHROME 7/16\" WITH WASHER (Box / 20) 15241","TYRE MACHINE INSERT - FLAT (EACH) BLACK","VERMAR BEAD SEALER 1L","EXTREME 5/150 CENTER CAPS / DOMES BLACK w/ WHITE WRITING","CHROME DOME MEDIUM 111X65MM (INCLUDES END CAP)","ADJUSTABLE WHEEL NUT INDICATOR 30mm-38mm (25) ORANGE","VERMAR 45MM SQUARE UNIVERSAL PATCHES SQAL45 (30/BOX)","GTS POWER BALANCE PB 1 (4/bag)","EXTREME CBL RING 106 - 100.1mm","TR418 - VALVE UNASSEMBLED (50)","METAL VALVE EXTENSION MEX100 (BRASS)","GTS EURO PASTE 1KG","VALVE CORE TOOL DOUBLE - RED","V3-14.2 TYRE VALVE (EACH)","STICK-ON Pb (25g*4) 30/BOX **BLACK**","EXTREME 6/139.7 CENTER CAPS / DOMES RED w/ BLACK WRITING","OPEN-ENDED 12X1.5MM WHEEL NUTS 19HEX (BOX 20) 11300","EXTREME 6/139.7 CENTER CAPS / DOMES BLACK w/ RED WRITING","EXTREME CBL RING 73.1 - 56.6mm","EXTREME CBL RING 73.1 - 63.4mm","EXTREME CBL RING 73.1 - 63.1mm","EXTREME CBL RING 73.1 - 70.2mm","EXTREME CBL RING 73.1 - 59.6mm","TR413 - VALVE COMPLETE (100)","EXTREME CBL RING 73.1 - 57.1mm","EXTREME CBL RING 73.1 - 65.1mm","EXTREME CBL RING 73.1 - 64.1mm","EXTREME CBL RING 73.1 - 56.1mm","EXTREME CBL RING 73.1 - 60.1mm","EXTREME CBL RING 73.1 - 69.6mm","EXTREME CBL RING 73.1 - 66.1mm","BRUSH - SWAB STYLE","EXTREME CBL RING 73.1 - 70.5mm","CHROME DOME SMALL 68mm TRAILER","SCREW-IN DOME 6/114.3& 6/130 66.1mm x 70mm BLACK (WITH LOGO)","PLASTIC VALVE CAP (100) RED","VERMAR INNER LINER SEALER 500ML","INTERNAL LARGE BORE REDUCER","STICK-ON Pb (50g*4) 20/box","EXTREME CBL RING 73.1 - 66.6mm","FLEXIBLE VALVE EXTENSION 160mm","EXTREME CBL RING 73.1 - 54.1mm","EXTREME CBL RING 110.3 - 100mm","BLACK DOME SMALL 68mm TRAILER","EXTERNAL LARGE BORE REDUCER","EXTREME CBL RING 73.1 - 67.1mm","7 SIDED HEX SOCKET K560","EXTREME CBL RING 110.3 - 106.1mm","SCREW-IN DOME 5/127 JEEP BLACK (WITH EXTREME LOGO)","4-WAY VALVE TOOL BLK PASSENGER","BLACK DOME LARGE 111X100MM (INCLUDES END CAP)","Wheel Bolt 14x1.5mm H17 S30 (30mm Thread) Black (20) 93730","Wheel Bolt 14x1.5mm H17 S28 (28mm Thread) Black (20) 93728","CHROME VALVE CAP \"DT\" SMVC9 (100)","CHROME DOME LARGE 111X100MM (INCLUDES END CAP)","PLASTIC VALVE CAP (100) BLACK","ACORN BLACK 1/2\" 45L LONG 21 HEX (Box of 20) 13940","ANGLED METAL VALVE EXT 90 DEG","VERMAR 50MM SQUARE UNIVERSAL PATCHES SQAL50 (30/BOX)","EXTREME 5/114.3 CENTER CAPS / DOMES BLACK w/ WHITE WRITING","BRUSH WITH PLASTIC HANDLE","ACORN BLACK 14X1.5mm 35L SHORT 19HEX (Box of 20) 13700","OPEN-ENDED 1/2\" WHEEL NUTS 19HEX (BOX 20) 11300","VS-8-90 BEADLOCKER VALVE (each","EXTREME CBL RING 110.3 - 67.1mm","WHEEL NUT INDICATOR 33MM (BAG OF 25) YELLOW","VERMAR 3MM COMBI PLUGS (25 PER BOX)","TUNER / INTERNAL HEX NUT 1/2\"  (Box of 20 + Key) RED","TR43E - TYRE VALVE (EACH)","METAL VALVE PULLER","WHEEL NUT INDICATOR 32MM (BAG OF 25) YELLOW","GTS TYRE VALVE DEFLATOR (POUCH WITH 4)","LOCKNUT SET 12X1.5 34mm (4+1 kEY) CHROME 43000","WHEEL NUT INDICATOR 21MM (BAG OF 25) YELLOW","BLACK DOME SMALL 75mm TRAILER","GTS EURO TYRE GREASE 5KG","Wheel Bolt 12x1.25mm H17 S30 (30mm Thread) Black (20) 93730","FAST FEEDER (NITTO)","BRUSH WITH WOODEN HANDLE 300MM","ANGLED METAL VALVE EXT 45 DEG","LOCKNUT SET 14X1.5 45L (4+1 KEY) CHROME","TR418 - VALVE COMPLETE (50)","PLASTIC CBL RING 108 - 93.2mm","GTS BLACK MOUNTING PASTE **1KG","PCD GAUGE (EXTREME)","LOCKNUT SET 12X1.5 34mm (4+1 KEY) BLACK 43000","ACORN BLACK 14X1.5mm 45L LONG 21HEX (Box of 20) 13940","WHEEL NUT INDICATOR 19MM (BAG OF 25) YELLOW","TYRE INFLATOR W/ PENCIL GAUGE (ATI06)","SPLINE LOCK SOCKET - STANDARD 6 SPLINE SLOTTED K596","Wheel Bolt 12x1.5mm H17 S30 (30mm Thread) Black (20) 93730","EXTREME CBL RING 110.3 - 105.1mm","BLACK DOME MEDIUM 111X65MM (INCLUDES END CAP)","STITCHER 50X6mm WOODEN HANDLE (50mm diameter)","TG4-07 BIAS PLY 178X178MM each","Wheel Bolt 12x1.25mm H17 S28 (28mm Thread) Black (20) 93728","EXTREME CBL RING 110.3 - 108.1mm","EXTREME CBL RING 110.3 - 78.1mm","SCREW-IN DOME 6/139.7 108mm x 70mm CHROME (W/ EXTREME LOGO)","4-WAY VALVE TOOL CHROME (VT07)","GTS TRUCK TYRE LEVER (FORCE)","PLASTIC VALVE EXTENSION EX150 (WHITE)","VALVE FISHING TOOL 300MM","GTS ATV REPAIR KIT Q644 - 25pcs (RED BOX)","TR546 - TYRE VALVE (EACH)","BLACK DOME 108X92MM (INCLUDES END CAP)","VERMAR XR-10 RADIAL PATCHES 55X75MM (45/BOX)","OE ALLOY WHEEL NUT CHROME 12x1.5mm w/ WASHER (24) 17200","GTS EURO PASTE 5KG (NEW)","ACORN CHROME 1/2\" 45L LONG 21HEX (Box of 20) 13940","BONDHUS L Wrench - Hex Short (HLS) 7/64 (Imperial)","BLACK DOME SMALL 84CB TRAILER","DIGITAL GAUGE STANDARD INFLATOR (ATI05)","VALVE FISHING TOOL 600MM","INVISIBLE VALVES (SET OF 4)","SPLINE NUT SET CHROME 14x1.5mm (Box of 20 + Socket) 49334","EXTREME 14X6 5/108 15P BLACK 75CB (HT FITMENT)","GTS BLACK MOUNTING PASTE 5KG","ACORN CHROME 9/16\" 45L 21HEX (Box of 20)","ACORN BLACK 14X1.5mm 45L LONG 21HEX (Box of 24) 13940","VALVE INSTALL TOOL GREEN","JOHNSON LEVEL I Beam 1200mm Structo-Cast  GLO Lime","EXTREME 13X4.5 MULTI 0P BLACK 75CB (BOX) (5/108/114.3)","TR544D - TYRE VALVE (EACH)","TUNER NUT 1/2\" CHROME","SCREW-IN DOME 6/114.3& 6/130 66.1mm x70mm CHROME (WITH LOGO)","OE ALLOY WHEEL NUT BLACK 14x1.5mm w/ WASHER (20) 17240","GTS SUPER HAND CLEANER 10L","Johnson Pocket Level w/ Strong Magnetic Pickup Tip & V-Groov","SPLINE NUT SET BLACK 12x1.5mm (Box of 20 + Socket) 49034","METAL VALVE EXTENSION MEX34 (CHROME)","Wheel Bolt 12x1.5mm H17 S40 (40mm Thread) Black (20) 93740","Wheel Bolt 12x1.5mm H17 S28 (28mm Thread) Black (20) 93728","SCREW-IN DOME 5/120 AMAROK BLACK (WITH EXTREME LOGO)","GTS SUPER HAND CLEANER 5L","ACORN PURPLE 12X1.5mm 35L SHORT 19HEX (Box of 24) 13700","ADJUSTABLE WHEEL NUT INDICATOR 24mm-30mm (25) ORANGE","ACORN BLACK 12X1.5mm 45L LONG 19HEX (Box of 20) 13745","ACORN BLACK 9/16\" 45L 21HEX (Box of 20) 13940","OPEN-ENDED 14x1.5mm WHEEL NUTS 7/8\" HEX (BOX OF 20)","OE ALLOY WHEEL NUT BLACK 14x1.5mm w/ WASHER (24) 17240","TUNER / INTERNAL HEX NUT 1/2\"  (Box of 20 + Key) BLACK 21032","EXTREME 5 SLOT 14X6J 5/114.3 6P SILVER 62cb","VERMAR 40MM SQUARE UNIVERSAL PATCHES SQAL40 (50/BOX)","ACORN BLACK 7/16\" - 35L 19HEX (Box of 20) 13700","TYRE MACHINE PLASTIC INSERT BEISBARTH / BOSCH (PACK OF 5)","LOCKNUT SET 7/16\" 32mm (4+1 KEY) CHROME 43000","CHEP PALLET","CP3/4\" DEEP 30MM IMPACT SOCKET","ACORN CHROME 14X1.5mm 35L SHORT 19HEX (Box of 20) 13700","TYRE LEVER SLEEVE - BLACK","NITTO HI CUPLA 1/4\" BSP 20PF FEMALE THREAD ADAPTOR","MAG NUT CHROME 1/2\" WITH WASHER (Box / 20) 15241","WHEEL NUT INDICATOR 21MM (BAG OF 50) YELLOW","CARBIDE CUTTER 8MM","EXTREME 16x7 6/139.7 10P WHITE 106.1CB","T-HANDLE KIT (CHROME)","LOCKNUT SET 1/2\" 34mm (4+1 KEY) BLACK 43000","ACORN BLACK 12X1.5mm 45L LONG 19HEX (Box of 24) 13745","GTS BOOTMAKER KNIFE W/ WOODEN HANDLE","VERMAR XR-12 RADIAL PATCHES 70X115MM (10/BOX)","CHROME DOME SMALL 84CB TRAILER","GTS TRUCK TYRE LEVER HALTE","SPIRAL PROBE 5\"","VALVE CORE TOOL LONG (RED)","TUBELESS VALVE STEM PULLER (RED WITH ALUMINIUM HANDLE)","ACORN PURPLE 14X1.5mm 35L SHORT 19HEX (Box of 20) 13700","CHROME DOME 108X92MM (INCLUDES END CAP)","TUNER / INTERNAL HEX NUT 12X1.25mm (Box of 20 + Key) RED","VALVE EXTENSION CLAMP - DOUBLE","VALVE FISHING TOOL 215MM","EXTREME 16x8-3 6/139.7 20P BLACK 106.1CB D-FIT DHOLE","VERMAR 6MM COMBI PLUGS (25 PER BOX)","SPLINE NUT SET BLACK 14x1.5mm (Box of 20 + Socket) 49034 NEW","JOHNSON LEVEL I Beam Structo-Cast 600mm","Metabo H1 Futuro Plus Keyless Chuck R+L 10mm w/ Plastic Slee","ACORN BLACK 12X1.5mm 35L 19HEX (Box of 24) 13700","EXTREME 14X6 5/112 30P WHITE D-HOLE 65CB","EXTREME 16x7-3 5/114.3 20P BLACK TRIANGLE 73.1CB","OE ALLOY WHEEL NUT BLACK 12x1.5mm w/ WASHER (24) 17200","ACORN CHROME 7/16\" 35L 19HEX (Box of 20) 13700","WHEEL NUT INDICATOR 35MM (BAG OF 25) YELLOW","TUNER / INTERNAL HEX NUT 12X1.5mm  (Box of 20 + Key) RED","1/4\" DIE GRINDER - GTS","OE ALLOY WHEEL NUT CHROME 14x1.5mm w/ WASHER (24) 17240","WHEEL NUT INDICATOR 34MM (BAG OF 25) YELLOW","ACORN PURPLE 14X1.5mm 45L LONG 21HEX (Box of 20) 13940","WHEEL NUT INDICATOR 38MM (BAG OF 25) YELLOW","OE ALLOY WHEEL NUT CHROME 14x1.5mm w/ WASHER (20) 17240","1 x Eclipse 300 x 12.5 x 0.025mm Hacksaw Blade 32TPI Bi Meta","ALPEN SPRINT MASTER DRILL 8.0 x 117mm (EACH)","ALPEN SPRINT MASTER DRILL 6.0 x 93mm (EACH)","VERMAR AL-2 ROUND TUBE PATCHES 53MM (70/BOX)","CBL ALUMINIUM 72.6mm-65.1mm (PACK OF 4)","SPLINE NUT SET BLACK 1/2\" (Box of 20 + Socket) 49034","SPLINE NUT SET CHROME 7/16\" (Box of 20 + Socket) 49034","ALPEN SPRINT MASTER DRILL 4.0 x 75mm (EACH)","EXTREME 16x8-3 6/139.7 0P BLACK 106.1CB TRIANGLE","EXTREME 15X7-2 6/139.7 10P WHITE 106.1CB","ACORN RED 1/2\" 35L SHORT 19HEX (Box of 20) 13700","ACORN PURPLE 12X1.25mm 35L 19HEX (Box of 24) 13700","RIM PROTECTOR (BLUE)","WHEEL WEIGHT STEEL 5G  (100)","ALPEN SPRINT MASTER DRILL 5.0 x 86mm (EACH)","ACORN RED 14X1.5mm 45L LONG 21HEX (Box of 20) 13940","OE ALLOY WHEEL NUT CHROME 12x1.25mm w/ WASHER (24) 17200","CP3/4\" DEEP 41MM IMPACT SOCKET","MAG NUT CHROME 12x1.5mm WITH WASHER (Box / 24) 15241","CARBIDE CUTTER 6MM","STICK-ON Fe(5G+10G) (50) BOX BLACK","JOHNSON CRAYON Lumber BLUE - Box Of 12","ALPEN SPRINT MASTER DRILL 3.0 x 61mm (EACH)","EXTREME 14X6 5/114.3 14P WHITE 84CB (PALLET) (FORD FITMENT)","AUTOSOL METAL POLISH 75ml TUBE #1000","EXTREME 12 SLOT 15X7J 5/114.3 0P SILVER 71.8cb","VERMAR 8MM COMBI PLUGS (25 PER BOX)","EXTREME 13X4.5 MULTI 0P GALV 75CB (PALLET) (5/108/114.3)","SPLINE NUT SET BLACK 12x1.25mm (Box of 24 + Socket) 49034","EXTREME CBL RING 110.3 - 93.2mm","REAMER / RASP 4\" WITH PLASTIC T-HANDLE GRIP","OPEN-ENDED 7/16\" WHEEL NUTS 19HEX (BOX 20) 11300","MARKAL \"B\" PAINTSTICK ORANGE (12)","Bondhus 5/32\" L Wrench Hex Short Singles (HLS) - ProGuard Fi","PCL ALLOY INFLATOR","VERMAR WHITE TYRE MARKING CRAYON (3/BLISTER PACK)","EXTREME 16x8 6/139.7 23N BLACK 110.1CB TRIANGLE","DIAL GAUGE STANDARD INFLATOR (ATI02)","EXTREME 16x8-3 6/139.7 20P BLACK 106.1CB TRIANGLE","ATV KIT REPLACEMENT NEEDLES (5) 120mm","ARBOUR 3/8\" - 6mm BUFFING TOOL 38.1mm LENGTH (SHORT)","ACORN BLACK 12X1.25mm 35L 19HEX (Box of 20) 13700","WHEEL NUT INDICATOR 27MM (BAG OF 25) YELLOW","CBL ALUMINIUM 72.6mm-66.9mm (PACK OF 4)","TYRE INFLATOR TDR DIGITAL *NEW","MARKAL \"B\" PAINTSTICK WHITE (12)","1 x Stanley 32TPI x 12\" Hacksaw Blade High Carbon Steel Flex","SPLINE NUT SET CHROME 12x1.5mm (Box of 24 + Socket) 49034","Wheel Bolt 12x1.5mm H17 S28 (28mm Thread) Chrome (20) 93728","CBL ALUMINIUM 66.1mm-57.1mm (PACK OF 4)","TUNER / INTERNAL HEX NUT 1/2\"  (Box of 20 + Key) CHROME","COMMAND AIR BLOW GUN 10mm 171.034","CBL ALUMINIUM 72.6mm-69.6mm (PACK OF 4)","CONTOUR 140 LOW PROFILE 31.7mm","VERMAR 63MM ROUND UNIVERSAL PATCHES XP-1 (15/BOX)","WHEEL NUT INDICATOR 22MM (BAG OF 25) YELLOW","EXTREME 16x7-3 5/114.3 20P BLACK D-FIT 73.1CB","GENERAL Lighted Tweezer 159mm Serated Blunt","STICK-ON STEEL (5G+10G)*4 (50) BOX","EXTREME IMITATION 17X8 5/127 6P BLACK SOFT8-LOCKER 71.5","TUBE PATCH ROUND TG30MM (98)","ALPEN SPRINT MASTER DRILL 4.5 x 80mm (EACH)","NITTO HI CUPLA 3/8\" BSP 30PH HOSE PLUG BARB TAIL ADAPTOR","ALPEN SPRINT MASTER DRILL 5.5 x 93mm (EACH)","LOCKNUT SET 12X1.5 34mm (4+1 KEY) BLUE 43000","VITOUR 185/70R14 88H CRUISER K365 WSW (20mm WHITEWALL)","NITTO HI CUPLA 3/8\" BSP 30SH HOSE BARB - SOCKET TYPE","CONTOUR 201 2\" NEW STYLE","NITTO HI CUPLA 3/8\" BSP 30SF FEMALE THREAD - SOCKET TYPE","ACORN GOLD 12X1.5mm 35L 19HEX (Box of 24) 13700","LOCKNUT SET 1/2\"\" 34MM (4+1 KEY) CHROME 43000","EXTREME 15X7-2 6/139.7 10P BLACK 106.1CB","ALPEN SPRINT MASTER DRILL 2.0 x 49mm (EACH)","Bondhus 1/8\" L Wrench Hex End Long Singles (HLS) - ProGuard","JOHNSON LEVEL Box 12000mm STD 3 VIAL Aluminium","ACORN BLACK 1/2\" 35L SHORT 21HEX (Box of 20) 13900","50G TUBELESS TRUCK TYRE WEIGHT (10)","WHEEL WEIGHT ALLOY 15G (100)","NITTO HI CUPLA 1/2\" BSPT 40SM MALE THREAD - SOCKET TYPE","NITTO HI CUPLA 3/8\" BSPT 30PM MALE THREAD PLUG","WHEEL WEIGHT STEEL 35G  (100)","CP3/4\" DEEP 27MM IMPACT SOCKET","WHEEL WEIGHT ALLOY MIXED","WHEEL WEIGHT STEEL 45G (40) **NEW BAGS OF 50**","LOCKNUT SET 1/2\" 34mm (4+1 KEY) BLUE 43000","ACORN GOLD 14X1.5mm 45L LONG 21HEX (Box of 20) 13940","1 x Eclipse 300 x 12.5 x 0.025mm Hacksaw Blade 24TPI HSS Har","NITTO HI CUPLA 1/2\" BSP 40SF FEMALE THREAD - SOCKET TYPE","MAG NUT CHROME 12x1.5mm WITH WASHER (Box / 20) 15241","LOCKNUT SET 14X1.5 35MM (4+1 KEY) BLACK 43035","CP AIR TOOL OIL 1L (Code CPA1014)","ALPEN SPRINT MASTER DRILL 3.5 x 70mm (EACH)","ACORN BLUE 12X1.25mm 35L 19HEX (Box of 24)","TYRE & TUBE MOUNTING COMP 10KG","CP3/4\" DEEP 36MM IMPACT SOCKET","Wheel Bolt 12x1.25mm H17 S30 (30mm Thread) Chrome (20) 93730","Metabo S1 Futuro Plus Keyless Chucks 10mm R+L w Plastic Sing","WHEEL WEIGHT STEEL MIXED","GTS WHEEL WEIGHT PLIERS (BLUE CASE)","75G TUBELESS TRUCK WEIGHT (10)","EXTREME 14X6 5/108 15P WHITE 75CB (HT FITMENT)","WHEEL WEIGHT STEEL 20G (100)","TUNER / INTERNAL HEX NUT 12X1.5mm  (Box of 20 + Key) BLUE","Sidchrome Socket 3/8\"DR 13/16\" 12PT (Imperial - Not Impact)","EXTREME 12 SLOT 14X6J 5/114.3 0P SILVER 71.8cb","Bondhus 2mm L Wrench Hex End Long Singles (HLL) - ProGuard F","EXTREME BNS 16X8 5/150 0P BLACK MACHINE FACE 110.5mm CB","EXTREME 5 SLOT 14X5J 5/114.3 6P SILVER 62cb","VERMAR 53MM ROUND UNIVERSAL PATCHES XP-0 (30/BOX)","TUNER / INTERNAL HEX NUT 12X1.25mm (Box of 20 + Key) BLUE","GTS TYRE PROBE (PLASTIC HANDLE","TREAD DEPTH GAUGE (PIRELLI)","WHEEL WEIGHT ALLOY 25G (100)","CP3/4\" DEEP 35MM IMPACT SOCKET","SPLINE NUT SET BLACK 12x1.5mm (Box of 24 + Socket) 49034","WHEEL WEIGHT STEEL 25G (100)","NITTO HI CUPLA 1/2\" BSP 40SH HOSE BARB - SOCKET TYPE","ACORN GOLD 1/2\" 34.5mm 35L SHORT 19HEX (Box of 20) 13700","ACORN RED 12X1.5mm 35L SHORT 19HEX (Box of 24) 13700","ADJUSTABLE WHEEL NUT INDICATOR 18mm-24mm (25) ORANGE","ALPEN SPRINT MASTER DRILL 4.1 x 75mm (EACH)","CP3/4\" DEEP 38MM IMPACT SOCKET","WHEEL WEIGHT ALLOY 10G (100)","RALLY 15X8 10/114.3/120.65 12N CHROME 81cb","ALPEN SPRINT MASTER DRILL 3.3 x 65mm (EACH)","JOHNSON CRAYON Lumber Yellow - Box Of 12","Wheel Bolt 12x1.25mm H17 S28 (28mm Thread) Chrome (20) 93728","ALPEN SPRINT MASTER DRILL 3.1x 62mm (EACH)","ALPEN SPRINT MASTER DRILL 4.8 x 81mm (EACH)","TR545D - TYRE VALVE (EACH)","KEN-TOOL VALVE CAPPER PRO","CP3/4\" DEEP 32MM IMPACT SOCKET","ARBOUR 3/8\" - 6mm BUFFING TOOL 69.8mm LENGTH (LONG)","NITTO HI CUPLA 3/8\" BSP 30PF FEMALE THREAD ADAPTOR","JOHNSON CRAYON Lumber RED - Box Of 12","SMOOTHIE 15X6 5/108 6P GLOSS BLACK 65.1cb","Wheel Bolt 12x1.5mm H17 S30 (30mm Thread) Chrome (20) 93730","ACORN BLUE 12X1.5mm 35L 19HEX (Box of 24)","ALPEN SPRINT MASTER DRILL 9.0 x 125mm (EACH)","GTS SUPER HAND CLEANER 1L *NEW","JOHNSON CRAYON Lumber BLACK - Box Of 12","ALPEN SPRINT MASTER DRILL 7.0 x 109mm (EACH)","ALPEN SPRINT MASTER DRILL 6.8 x 109mm (EACH)","MAG NUT BLACK 1/2\" WITH WASHER (Box / 20) 15241","ALPEN SPRINT MASTER DRILL 1.5 x 40mm (EACH)","ALPEN SPRINT MASTER DRILL 1.0 x 34mm (EACH)","WHEEL WEIGHT ALLOY 5G  (100)","JAMEC PEM DIGITAL AUTOMATIC WALL INFLATOR","LOCKNUT SET 12X1.25 34mm (4+1 KEY) BLUE 43000","Wheel Bolt 14x1.25mm H17 S30 (30mm Thread) Black (20) 93730","ALPEN SPRINT MASTER DRILL 8.5 x 117mm (EACH)","TR543D - TYRE VALVE (EACH)","EXTREME 16x7 **5/139.7** 13N BLACK 110.1CB TRIANGLE","ALPEN SPRINT MASTER DRILL 4.2 x 75mm (EACH)","ALPEN SPRINT MASTER DRILL 7.5 x 109mm (EACH)","EXTREME 15X7 6/139.7 13N BLACK 111CB","CHROME SMOOTHIE CAP (LOOSE)","Bondhus 1/4\" L Wrench Hex Short Singles (HLS) - ProGuard Fin","EXTREME 16x8-3 6/139.7 0P BLACK 110.1CB D-FIT DHOLE","TUNER / INTERNAL HEX NUT 1/2\"  (Box of 20 +Key) PURPLE 21032","TUNER / INTERNAL HEX NUT 12X1.5mm (20 + Key) PURPLE 21032","EXTREME 14X6 5/114.3 0P BLACK 84CB (FORD FITMENT)","MARKAL \"B\" PAINTSTICK RED (12)","LOCKNUT SET 7/16\" 32mm (4+1 KEY) BLACK 43000","GTS PREMIUM TYRE LEVER 300MM","ALPEN SPRINT MASTER DRILL 10.0 x 133mm (EACH)","ACORN BLUE 14X1.5mm 45L LONG 21HEX (Box of 20) 13940","GTS PREMIUM TYRE LEVER 600MM","RUBBER PAD PASSENGER CAR LIFT - 160 X 120 X 40MM","SURETORQ 1.6X8.0 150MM SLOTTED (FLAT) SCREWDRIVER 09660744","LOCKNUT SET 12X1.25 34MM (4+1 KEY) CHROME 43000","CP3/4\" 17MM SQ DRIVE SOCKET","SURETORQ 1.0X5.5 75MM SLOTTED (FLAT) SCREWDRIVER 09670740","WHEEL WEIGHT ALLOY 20G (100)","GENERAL PickTool Lighted Magnetic Telescopic","ADJUSTABLE WHEEL NUT INDICATOR 38mm-50mm (25) ORANGE","150G TUBELESS TRUCK WEIGHT (10)","TUNER / INTERNAL HEX NUT 12X1.25mm (20 + Key) PURPLE 21032","Sidchrome Socket 1/4\"DR 5mm 6PT DEEP (Metric - Not Impact)","N1091 PUSH BUTTON WATER ADAPTOR","WHEEL WEIGHT ALLOY 35G (100)","SURETORQ PH#1 75MM PHIILIP SCREWDRIVER 09670101","SURETORQ 1.6X8.0 200MM SLOTTED (FLAT) SCREWDRIVER 09680744","BONDHUS L Wrench - Hex Long (HLL) 3/32 (Imperial)","EXTREME 15X7 *5/139.7* 13N BLACK D-FIT 110.1CB","SURETORQ PH#3 200MM PHIILIP SCREWDRIVER 09680103","ACORN PURPLE 1/2\" 35L SHORT 19HEX  (Box of 20) 13700","SURETORQ 1.2X6.5 125MM SLOTTED (FLAT) SCREWDRIVER 09650742","ALPEN SPRINT MASTER DRILL 6.5 x 101mm (EACH)","ACORN GOLD 12X1.25mm 35L 19HEX (Box of 24) 13700","ALPEN SPRINT MASTER DRILL 2.5 x 57mm (EACH)","GTS HEAVY DUTY WEIGHT PLIERS (CLEAR PACKET RED HANDLE)","OPEN EYE NEEDLE 4\" (SINGLES)","MAG NUT BLACK 7/16\" WITH WASHER (Box / 20) 15241","SURETORQ PH#3 150MM PHIILIP SCREWDRIVER 09660103","Metabo 1-10mm Futuro Top Keyless Chuck R 1/2\"-20 UNF w/ Firm","SURETORQ PH#2 100MM PHIILIP SCREWDRIVER 09640102","TYRE MACHINE INSERT - PIN TYPE (EACH) YELLOW","COMMAND 8PC SLIMLINE SOLDERING IRON / BLOW TORCH","LOCKNUT SET 12X1.25 34mm (4+1 KEY) BLACK 43000","EXTREME 17x8-3 6/130 30P BLACK 75.6cb TRIANGLE","SURETORQ 1.2X6.5 100MM SLOTTED (FLAT) SCREWDRIVER 09640742","EXTREME 12 SLOT 14X7J 5/114.3 0P CHROME 71.8cb","JOSCO STRIP IT DISC 100mm RED 16mm BORE","ACORN CHROME 12X1.5mm 45L LONG 19HEX (Box of 24) 13745","GTS SPLIT EYE TOOL","EXTREME 17x8-3 6/130 30P BLACK 75.6cb R-FIT ROUND","ALLOY 8.25R22.5 10X335 10x26mm  POLISHED INSIDE","Wheel Bolt 12x1.5mm H17 S40 (40mm Thread) Chrome (20) 93740","ALLOY RIM 8.25R22.5 10H MACHINE","Sidchrome Socket 3/8\"DR 5mm 6PT (Metric - Not Impact)","TYRE MACHINE PLASTIC INSERT CORGHI / FAIP (PACK OF 5)","WHEEL WEIGHT ALLOY 45G (50) **NEW BAGS OF 50**","SURETORQ PH#2 125MM PHIILIP SCREWDRIVER 09650102","ACORN BLUE 12X1.25mm 35L 19HEX (Box of 20) 13700","Sidchrome Socket 1/4\"DR 9/32\" 6PT DEEP","WHEEL WEIGHT FN 45G (100)","Sidchrome Socket 3/8\"DR 7/8\" 12PT (Imperial - Not Impact)","Sidchrome Socket 3/8\"DR 9mm 12PT (Metric - Not Impact)","ACORN CHROME 12X1.25mm 35L 19HEX (Box of 24) 13700","4-WAY VALVE TOOL BLK E/MOVER","JOHNSON LEVEL Box 600mm STD 3 VIAL Aluminium","BONDHUS L Wrench - Hex Long (HLL) 1/16 (Imperial)","CP3/4\" 19MM SQ DRIVE SOCKET","EXTREME 16x8 6/139.7 13N BLACK 110.1CB D-FIT D-HOLE","EXTREME 17x8-3 6/139.7 23N BLACK 110.1CB TRIANGLE 1400KG","SPLINE NUT SET BLACK 7/16\" (Box of 20 + Socket) 49034","EXTREME 17x8-3 6/130 30P BLACK 75.6cb D-FIT","METAL VALVE EXTENSION MEX110 (CHROME)","BONDHUS L Wrench - Hex Long 3/16 (Imperial)","WHEEL WEIGHT STEEL 30G (100)","EXTREME 12 SLOT 15X8J 5/114.3 19N SILVER 71.8cb","Bondhus 9/64\" L Wrench Ball Hex End Long (BL) - ProGuard Fin","Sidchrome Socket 1/2\"DR 3/8\" 12PT (Imperial - Not Impact)","CBL ALUMINIUM 74.1mm-72.6mm (PACK OF 4)","TVS 12-16.5 ST30 12PR TL  R-4","NITTO HI CUPLA 1/2\" BSP 40PH HOSE PLUG BARB TAIL ADAPTOR","Sidchrome Socket 3/8\"DR 11/16\" 12PT (Imperial - Not Impact)","RUBBER PAD PASSENGER CAR LIFT - 160 X 120 X 60MM","EXTREME 16x7-3 6/114.3 35P BLACK R-FIT 66.1CB ROUND","Sidchrome SPANNER RING & OPEN END 11/16\"","BONDHUS L Wrench - Ball End Long (BL) 5/64 (Imperial)","TYRE INFLATOR TDR3000 (NEW)","GENERAL ScrewDriver Set Lighted With 3 Rev.Blades","TVS 10-16.5 ST30 10PR TL  R-4","TUNER / INTERNAL HEX NUT 12X1.25mm (20 + Key) BLACK 21032","8PC 1/2\" Dr Thin Wall Lug Wheel Nut Remover Deep","12-22MM STANLESS BAND WORM DRIVE CLAMP","Sidchrome Socket 1/4\"DR 3/16\" 6PT (Imperial - Not Impact)","Sidchrome Socket 1/2\"DR 9mm 12PT (Metric - Not Impact)","EXTREME IMITATION 17X8 6/139.7 13N BLACK D-LOCKER 110.1","Sidchrome Socket 1/2\"DR 1/2\" 12PT DEEP","EXTREME 14X6 MULTI 0P GALV 84CB (BOX) (5/108/114.3)","Sidchrome Socket 1/2\"DR 3/8\" 12PT DEEP","LOCKNUT SET 14X1.5 45L (4+1 KEY) BLACK","Sidchrome INHEX SOCKET","7PC 1/2\" Dr Wheel Nut Deep Impact / Flip SKT Set","BONDHUS L Wrench - Hex Short (HLS) 1/16 (Imperial)","EXTREME 16x7 6/139.7 13N BLACK 110.1CB","EXTREME 16x7 6/139.7 10P BLACK 106.1CB","SMOOTHIE 15X6 5/108 6P CHROME 65.1cb","4 X SIDED WHEEL BRACE (17X19X21X22)","TYRE MACHINE PLASTIC INSERT HOFFMAN (Pack of 5)","NITTO HI CUPLA 1/4\" BSPT 20SM MALE THREAD - SOCKET TYPE","WHEEL WEIGHT FN 5G (100)","UNI PATCH ROUND TGU30MM (80)","BONDHUS L Wrench - Hex Short (HLS) 7/16 (Imperial)","EXTREME 12 SLOT 14X5J 5/114.3 0P SILVER 71.8cb","Sidchrome Socket 1/4\"DR 3/16\" 6PT DEEP","Sidchrome Socket 1/4\"DR 7/16\" 6PT DEEP","EXTREME OEM STYLE 15X7J 5/120.65 0P CHROME 71.3cb","NITTO HI CUPLA 1/4\" BSP 20SH HOSE BARB - SOCKET TYPE","RIM PROTECTOR (YELLOW)","WHEEL WEIGHT ALLOY 60G (50) **NEW BAGS OF 50**","Sidchrome Socket 3/8\"DR 7/16\" 12PT (Imperial - Not Impact)","EXTREME 12 SLOT 14X7J 5/114.3 0P SILVER 71.8cb","BONDHUS L Wrench - Hex Long (HLL) 7/64 (Imperial)","GTS PREMIUM TYRE LEVER 500MM","GENERAL Thermo - Hygrometer Pen Digital Display","GENERAL Lighted Tweezer 159mm Serated Bent","EXTREME 16x8 6/139.7 0P BLACK 110.1CB TRIANGLE","WHEEL WEIGHT FN 25G (100)","EXTREME 16x7 **5/139.7** 13N BLACK 110.1CB D-FIT","NITTO HI CUPLA 1/4\" BSP 20SF FEMALE THREAD - SOCKET TYPE","Sidchrome Socket 3/8\"DR 3/8\" 12PT (Imperial - Not Impact)","WHEEL WEIGHT STEEL 10G (100)","WHEEL WEIGHT ALLOY 55G (100)","WHEEL WEIGHT STEEL 15G (100)","CHROME RALLYE CAP (LOOSE)","ADHESIVE WHEEL WEIGHT CUTTER (HANDHELD)","EXTREME IMITATION 17X9 5/127 0P BLACK SOFT8-LOCKER 71.5","RIM WIDTH CALIPER","Josco Brumby 10 Piece Multi Purpose Drill Accessory Kit BDAK","WELD STYLE MAG NUT CHROME 1/2\" WITH WASHER (Box / 20) 15260","Sidchrome Socket 1/2\"DR 7/8\" 12PT (Imperial - Not Impact)","EXTREME 17x8 6/139.7 0P BLACK 106.1CB 1400KG D-FIT 1400kg","WHEEL WEIGHT FN 15G (100)","SMOOTHIE 15X10 5/108 25N GLOSS BLACK 65.1cb","LUG NUT CAP PULLER","VITOUR 225/95R16C (8) 118/116T DESERT ARK TUBE TYPE","WHEEL WEIGHT FN 55G (100)","EXTREME 16x7-3 5/120 25P BLACK 65.1CB AMAROK D-FIT","ACORN BLACK 1/2\" 35L SHORT 19HEX (Box of 20) 13700","WHEEL NUT INDICATOR 41MM (BAG OF 50) YELLOW","G/YEAR 255/70R22.5 140/137M G622 RSD H M&S","EXTREME OEM STYLE 15X8J 5/120.65 6P CHROME 71.3cb","EXTREME OEM STYLE 15X10J 5/120.65 25N CHROME 71.3cb","GTS RUBBER MALLET - 300MM","Sidchrome Socket 3/8\"DR 3/8\" 6PT DEEP","Bondhus 5/32\" L Wrench Ball End Long (BL) - ProGuard Finish","EXTREME 5 SLOT 14X7J 5/114.3 6P SILVER 62cb","STICK-ON ROLL Fe 6KG (5G) BLACK","JOSCO STRIP IT DISC 115mm RED 22mm BORE","SURETORQ Extractor 12mm x 1/4\" Dr Hex Shank","GENUINE BEADLOCK 17X8 6/139.7 23N D-HOLE BLACK 110.1cb","Sidchrome Socket 1/4\"DR 7/32\" 6PT DEEP","Bondhus 9/64\" L Wrench Singles Ball End Long - ProGuard Fini","SURETORQ EXTRACTOR SOCKET 3/8\" x 3/8 Dr","SURETORQ EXTRACTOR SOCKET 1/2\" x 3/8 Dr","Sidchrome Socket 1/4\"DR 6mm 6PT DEEP (Metric - Not Impact)","JAMEC PEM ACS1500 DIGITAL TYRE INFLATOR","EXTREME 17x8-3 6/139.7 30P BLACK 106.1CB D-FIT 1400KG DHOLE","EXTREME IMITATION 16X8 5/150 0P BLACK D-LOCKER 110.1cb","Sidchrome Socket 3/8\"DR 3/4\" 12PT (Imperial - Not Impact)","Sidchrome Socket 3/8\"DR 11/16\" 6PT DEEP","EXTREME 16x7-3 5/120 25P BLACK 65.1CB AMAROK TRIANGLE","Wheel Bolt 14x1.5mm H17 S30 (30mm Thread) Chrome (20) 93730","BONDHUS L Wrench - Ball End Long (BL) 3/32 (Imperial)","WHEEL WEIGHT FN 35G (100)","Milwaukee 30 Litres Plastic Dust Filter Bags Disposable - 5","Bondhus 3/8\" L Wrench Hex End Long Singles (HLL) - ProGuard","Sidchrome Socket 1/2\"DR 9/16\" 12PT DEEP","EXTREME IMITATION 16X8 5/150 25N BLACK D-LOCKER 110.1cb","EXTREME 16x8 5/120 35P BLACK 65.1CB TRIANGLE 1400kg","Sidchrome Socket 3/8\"DR 1/4\" 6PT (Imperial - Not Impact)","EXTREME 17x8-3 6/139.7 13N BLACK 110.1CB D-FIT 1400KG DHOLE","Bondhus 1/8\" L Wrench Hex Ball End Short - ProGuard Finish I","SURETORQ EXTRACTOR SOCKET 1/4\" x 3/8 Dr","3/4\" DR SOCKET RETAINER RING (PACK OF 5)","GENERAL Screwdriver Precision Set Powered 6X 3mm Bit","Paslode 38mm Timber Flooring Staples 80 Series Electro Galv","SURETORQ Extractor 6mm x 1/4\" Dr Hex Shank","22mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522022","Sidchrome Socket 1/4\"DR 12mm 6PT DEEP (Metric - Not Impact)","Sidchrome Socket 1/4\"DR 4mm 6PT (Metric - Not Impact)","EXTREME IMITATION 17X8 6/139.7 23N BLACK D-LOCKER 110.1","ABW / Sidchrome 7/16\" 1/2\"DR Deep Impact Socket - Imperial","EXTREME 15X7 5/114.3 0P WHITE 75CB","EXTREME 16x7-3 6/114.3 35P WHITE 66.1CB","EXTREME 15X8 5/139.7 23N BLACK D-FIT 111CB","Bondhus 9mm L Wrench Ball End Long (BL) - ProGuard Finish Me","PRONAR 16X22.5 WHEELS YELLOW","WHEEL WEIGHT STEEL 40G (100)","CARLISLE 9/3.50X4 CAREFREE SOLID SLICK","Bondhus 6mm L Wrench Hex End Long (HLL) - ProGuard Finish Me","ELIX HAND CLEANING WIPES (90 PER BUCKET)","EXTREME 17x8-3 6/139.7 23N BLACK 110.1CB D-FIT 1400KG DHOLE","Sidchrome Socket 1/2\"DR 11/16\" 12PT DEEP","55G L/T WEIGHT (25)","Sidchrome Socket 1/4\"DR 9/32\" 6PT (Imperial - Not Impact)","WHEEL WEIGHT FN 50G (50) **NEW BAGS OF 50**","EXTREME IMITATION 17X8 6/139.7 20P BLACK D-LOCKER 106.1","AUTOSOL \"SHINE\" METAL POLISH 50G TUBE #1187","Sidchrome Socket 3/8\"DR 5/16\" 12PT (Imperial - Not Impact)","Sidchrome Socket 1/2\"DR 7/16\" 12PT DEEP","Metabo 1.5 - 13mm Gear Chuck w/ Key for Clockwise Rotating D","MARKAL \"B\" PAINTSTICK BLUE (12)","SUMMIT LT305/70R18 10PR MUD HOG TL 126/123 Q","WHEEL WEIGHT FN MIXED","Bondhus 9/64\" L Wrench Hex Short Singles - ProGuard Finish I","EXTREME OEM STYLE 15X8J 5/120.65 25N CHROME 71.3cb","SURETORQ Extractor 5/16 (8mm) x 1/4\" Dr Hex Shank","EXTREME 17x8-3 6/139.7 45P BLACK 106.1CB R-FIT 1400KG ROUND","EXTREME 16x8-3 6/130 20P BLACK 75.6cb D-FIT","JOSCO BTC502 65mm X 10X1.5mm TWISTKNOT CUP BRUSH","Sidchrome Socket 3/8\"DR 7/16\" 6PT DEEP","RALLYE 15X10 10/114.3/120.65 25N CHROME 81cb","SURETORQ EXTRACTOR SOCKET 7/16\" (11mm) x 3/8 Dr","100G TUBELESS TRUCK WEIGHTS (10)","ABW / Sidchrome 11mm 1/2\"DR Standard Impact Socket - Metric","EXTREME 15X7 5/114.3 0P BLACK D-FIT 75CB","EXTREME IMITATION 17X9 6/139.7 23N BLACK D-LOCKER 110.1","Sidchrome Socket 1/2\"DR 5/8\" 12PT DEEP","SLIDE HAMMER 1170-1830MM","TUNER / INTERNAL HEX NUT 12X1.5mm ( (Box of 20 + Key) CHROME","SURETORQ Extractor 10mm x 1/4\" Dr Hex Shank","Sidchrome Socket 3/8\"DR 9/16\" 6PT DEEP","EXTREME 18X8 6/114.3 35P BLACK 66.1cb D-FIT","WHEEL WEIGHT TOYOTA STEEL MIX (50)","200G TUBELESS TRUCK WEIGHT (10)","EXTREME IMITATION 17X8 5/120 30P BLACK D-LOCKER 65.1 AMAROK","EXTREME 15X7 6/139.7 13N WHITE 111CB","JOSCO BTC503 65mm X 14X2mm TWISTKNOT CUP BRUSH","BONDHUS L Wrench - Hex Long (HLL) 5/64 (Imperial)","EXTREME 17x8-3 5/120 30P BLACK 65.1CB R-FIT 1400KG ROUND","NITTO HI CUPLA 1/4\" BSPT 20PM MALE THREAD PLUG","ABW / Sidchrome 11/16\" 1/2\"DR Deep Impact Socket - Imperial","EXTREME 16x8 5/120 35P BLACK 65.1CB D-HOLE 1400kg","12PC 1/2\" Dr Wheel Lug Nut Remover Impact SKT Set","SURETORQ EXTRACTOR SOCKET 5/8\" (16mm) x 3/8 Dr","Sidchrome Socket 3/8\"DR 3/4\" 6PT DEEP","EXTREME 16x8-3 6/139.7 0P BLACK 106.1CB D-FIT DHOLE","Sidchrome SPANNER RING & OPEN END 5/8\"","PCL 18PC 1/2\" IMPACT SOCKETSET","Sidchrome Socket 3/8\"DR 11mm 6PT DEEP (Metric - Not Impact)","WHEEL WEIGHT TOYOTA STEEL 50g (50)","WHEEL WEIGHT TOYOTA STEEL 70g (50)","Sidchrome Socket 1/4\"DR 3/8\" 6PT DEEP","HOSE REEL PRO  EXTREME AIR 20M 58.3012","ALPEN SPRINT MASTER DRILL 9.5 x 125mm (EACH)","PCL 27MM IMPACT SOCKET (3/4\" DRIVE)","DOT CODE CUTTING TOOL","Wheel Bolt 14x1.5mm H17 S40 (40mm Thread) Chrome (20) 93740","WHEEL WEIGHT ALLOY 30G (100)","RUBBER PAD PASSENGER CAR LIFT - 160 X 120 X 80MM","RUBBER PAD WHEEL CHOCK","TVS 11L-15SL (280L/70-15SL)  12PR TL  I-1","ABW / Sidchrome 12mm 1/2\"DR Deep Impact Socket - Metric","250G TUBELESS TRUCK WEIGHT (10)","EXTREME 17x8-3 6/139.7 45P BLACK 106.1CB 1400KG TRIANGLE","EXTREME 16x7 5/150 0P BLACK 110.1CB 1400KG","ABW / Sidchrome 16mm 1/2\"DR Deep Impact Socket - Metric","EXTREME 16x8-3 6/130 20P BLACK 75.6cb TRIANGLE","TVS 9.5L-15SL (240L/80-15SL) 12PR TL  I-1","TVS 12.5L-15SL (320L/70-15SL) 12PR TL  I-1","COMMAND ELECTRICAL RESIN CORE SOLDER","EXTREME 14X6 5/120.65 0P BLACK 84CB (HQ PATTERN)","Sidchrome Socket 1/2\"DR 11/16\" 12PT (Imperial - Not Impact)","SMOOTHIE 15X7 5/108 6P GLOSS BLACK 65.1cb","BONDHUS L Wrench - Ball End Long 3/16 (Imperial)","Sidchrome Socket 1/4\"DR 3/8\" 6PT (Imperial - Not Impact)","TVS 12.5L-16SL (320L/70-16SL) 10PR TL  I-1","EXTREME IMITATION 16X8-3 6/139.7 20P BLACK D-LOCKER 106.1","BONDHUS L Wrench - Hex Short (HLS) 5/16 (Imperial)","TVS 460/70R24 TLB504 159A8/159B TL  TAUROTRAC R4 SB","MARKAL \"B\" PAINTSTICK YELLOW (12)","BONDHUS L Wrench - Ball End Long (BL) 7/64 (Imperial)","Metabo B16 Futuro Keyless Chuck R 1-10mm for Clockwise Rotat","Sidchrome Socket 3/8\"DR 5/8\" 12PT (Imperial - Not Impact)","Sidchrome Socket 3/8\"DR 18mm 6PT DEEP (Metric - Not Impact)","Sidchrome Socket 1/4\"DR 7/32\" 6PT (Imperial - Not Impact)","EXTREME 17x8 6/139.7 0P BLACK 106.1CB 1400KG TRIANGLE 1400kg","Sidchrome Socket 3/8\"DR 1/2\" 12PT (Imperial - Not Impact)","Sidchrome SPANNER RING & OPEN END 7/8\"","Sidchrome Socket 1/4\"DR 11mm 6PT (Metric - Not Impact)","SURETORQ EXTRACTOR SOCKET 17mm x 3/8 Dr","Sidchrome Socket 1/4\"DR 1/2\" 6PT DEEP","110G L/T WEIGHT (25)","UNION TAPERED SPINDLE SET R/H 1/2\" & 12mm","EXTREME IMITATION 16X8 6/139.7 13N BLACK D-LOCKER 110.1cb","TRACTOR TYRE GAUGE","Sidchrome Socket 1/4\"DR 1/4\" 6PT DEEP","EXTREME IMITATION 15X8 *5/139.7* 23N BLACK D-LOCKER 111","Paslode 45mm Timber Flooring Staples 80 Series Electro Galv","TUNER / INTERNAL HEX NUT 12X1.5mm  (20 + Key) BLACK 21032","Sidchrome SPANNER RING & OPEN END 13/16\"","EXTREME IMITATION 16X8 5/120 35P BLACK D-LOCKER 65.1","ALLOY RIM 8.25R22.5 10H POLISHED","22mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124022B","Sidchrome Socket 1/2\"DR 7/16\" 12PT (Imperial - Not Impact)","Bondhus 7mm L Wrench Ball End Long (BL) - ProGuard Finish Me","EXTREME 16x8-3 6/139.7 0P BLACK 110.1CB R-FIT ROUND","EXTREME 18X8 6/139.7 0P BLACK 110.1cb D-FIT","4PC 3/8\" Dr Metric Deep Impact Wheel Nut SKT Set","Sidchrome Socket 1/4\"DR 11/32\" 6PT (Imperial - Not Impact)","Bondhus 5.5mm L Wrench Ball End Long (BL) - ProGuard Finish","Bondhus 1.5mm L Wrench Hex Short Singles (HLS) - ProGuard Fi","SMOOTHIE 15X7 10/114.3/120.65 6P GLOSS BLACK 81cb","Sidchrome Socket 1/4\"DR 4mm 6PT DEEP (Metric - Not Impact)","SMOOTHIE 15X7 5/108 6P CHROME 65.1cb","TVS 500/50-17(19.5/50-17) IM126 18PR TL  I-1","ALPEN PRO SPLIT POINT DRILL BITS KM25 (SET)","Bondhus 7/32\" L Wrench Hex Short Singles (HLS) - ProGuard Fi","SURETORQ EXTRACTOR SOCKET 3/4\" (19mm) x 3/8 Dr","Bondhus 1/4\" L Wrench Hex End Long (HLL) - ProGuard Finish I","Sidchrome Socket 1/4\"DR 5/16\" 6PT (Imperial - Not Impact)","Sidchrome Socket 1/2\"DR 11mm 12PT (Metric - Not Impact)","Sidchrome Socket 3/8\"DR 20mm 12PT (Metric - Not Impact)","Sidchrome Socket 3/8\"DR 1/2\" 6PT DEEP","Sidchrome Socket 3/4\"DR 30mm 12PT (Metric - Not Impact)","ALPEN SPRINT MASTER DRILL KM25 (SET)","Sidchrome Socket 1/2\"DR 16mm 12PT DEEP (Metric - Not Impact)","TRUCK LEVER KIT","COMMAND 4PC BUTANE SOLDERING TIP SET","WHEEL WEIGHT FN 40G (100)","Sidchrome Socket 1/4\"DR 7mm 6PT DEEP (Metric - Not Impact)","Sidchrome Socket 3/8\"DR 11mm 12PT (Metric - Not Impact)","7PC 1/2\" Dr Wheel Lug Nut Remover Deep Impact SKT Set","Bondhus 9/16\" L Wrench Ball End Long (BL) - ProGuard Finish","Sidchrome Socket 1/4\"DR 5/16\" 6PT DEEP","DUAL ACTION 5 PIECE ADAPTOR SET","JOURNEY 13X6.50-6 (6) P332","Sidchrome Socket 1/4\"DR 5mm 6PT (Metric - Not Impact)","COMMAND 22PC 240V 30W WOOD BURNING KIT","TUNER / INTERNAL HEX NUT 12X1.25mm (20 + Key) CHROME 21032","Sidchrome Socket 3/8\"DR 9/16\" 12PT (Imperial - Not Impact)","Bondhus 1/8\" L Wrench Hex End Short Singles (HLS) - ProGuard","WHEEL WEIGHT STEEL 50G (50) **NEW BAGS OF 50**","WHEEL WEIGHT FN 30G (100)","Makita E-12619 24pce Impact Driver Bit Set","Sidchrome Socket 3/4\"DR 1-3/16\" 12PT (Imperial - Not Impact)","Josco Brumby 25mm Shaft Mounted Calico Polish Buff BCC2525","21mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522021","Poly Flap Disc 115x22 Coarse","SURETORQ 27PCS MINI RATCHET AND BIT SET","SUMMIT LT295/60R20 10PR MUD HOG TL 126/123 Q","BONDHUS L Wrench - Ball End Long (BL) 5/16 (Imperial)","Xcelite 5/32\" x 6\" Allen Hex End Screwdriver for Recessed So","EXTREME 17x8-3 6/139.7 45P BLACK 106.1CB D-FIT 1400KG D-HOLE","ABW / Sidchrome 7/8\" 1/2\"DR Deep Impact Socket - Imperial","ABW / Sidchrome 20mm 1/2\"DR Standard Impact Socket - Metric","Sidchrome Socket 1/4\"DR 1/2\" 6PT (Imperial - Not Impact)","EXTREME 14X6 6/139.7 0P 110.5CB 5603 ALLOY","EXTREME 18X8 5/120 30P BLACK 65.1cb D-FIT","ABW / Sidchrome 10mm 1/2\"DR Standard Impact Socket - Metric","TUNER NUT KEY (K468)","Poly Flap Disc 115x22 Medium","RALLY 15X6 10/114.3/120.65 6P CHROME 81cb","TVS 7.50-16 TF27 8PR TT  F-2","Kincrome SPANNER D/R 1/4\" - 5/16\"","EXTREME 17x8-3 6/114.3 35P BLACK 66.1CB FIT 1400KG TRIANGLE","300G TUBELESS TRUCK WEIGHT (10)","TVS 6.00-16 TF27 8PR TT  F-2","EXTREME 15X8 5/139.7 23N BLACK 111CB","UNION TAPERED SPINDLE SET L/H 1/2\" & 12mm","ABW / Sidchrome 12mm 1/2\"DR Standard Impact Socket - Metric","Sidchrome Socket 1/4\"DR 7/16\" 6PT (Imperial - Not Impact)","GTS 3MM COMBI REPAIR (60)","TYRE LEVER - 400mm FOR MOTORBIKE / SCOOTER","WHEEL WEIGHT TOYOTA STEEL 30g (50)","TVS 7.50-15 CT09 14PR TT  C-1","EXTREME 12 SLOT 15X7J 5/114.3 0P CHROME 71.8cb","SMOOTHIE 15X10 10/114.3/120.65 25N CHROME 81cb","SURETORQ EXTRACTOR SOCKET 5/16\" (8mm) x 3/8 Dr","EXTREME 16x8 5/150 25N BLACK 110.1CB TRIANGLE","Sidchrome Socket 3/8\"DR 6mm 6PT (Metric - Not Impact)","Bondhus 3.5mm L Wrench Hex Short Singles (HLS) - ProGuard Fi","Sidchrome Socket 1/4\"DR 11mm 6PT DEEP (Metric - Not Impact)","Sidchrome Socket 1/2\"DR 1-1/16\" 12PT (Imperial - Not Impact)","Johnson Cross Check Level Tool - High-Impact Body w/ Mountin","Sidchrome Socket 1/2\"DR 3/4\" 12PT DEEP","Sidchrome Socket 3/4\"DR 7/8\" 12PT (Imperial - Not Impact)","Makita E-12631 30pce Impact Driver Bit Set","Bondhus 7/32 \" L Wrench Hex End Long (HLL) - ProGuard Finish","UNION WIRE BRUSH E/B 0.35 T/K 20mm RD/SK DB","Bondhus 9/16\" L Wrench Hex End Long Singles (HLL) - ProGuard","EXTREME IMITATION 16X8 6/139.7 23N BLACK D-LOCKER 110.1cb","TUBE PATCH ROUND TG55MM (40)","SMOOTHIE 15X10 5/108 25N CHROME 65.1cb","BONDHUS L Wrench - Hex Long (HLL) 5/16 (Imperial)","CP STUD CLEANERS (SET OF 4)","Airco 45mm Brad Nail C Series Electro Galvanise Plain Shank-","WHEEL WEIGHT FN 60G (50) **NEW BAGS OF 50**","KINCROME 3/8\" 1/2\"DR HEX IMPACT SOCKET","TVS 15.5/80-24(400/80-24) IM54 HD 16PR TL  R-4","SAMSON 7.50-16 14PR SAv C-1/C-1A Set 14PR","ABW 1-1/16 1/2 DRIVE IMPACT SOCKET","Josco Brush Cup CR HS 25 6.3 Spindle 0.33","BKT 20x12-10 (4) LG307","21mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024021B","WHEEL WEIGHT FN 10G (100)","EXTREME 16x7-3 6/114.3 35P BLACK D-FIT 66.1CB","WHEEL WEIGHT STEEL 70G (50)","ABW ADAPTOR 3/4\"DR TO 1\"DR","KOKEN SOCKET DEEP 1/2DR X 1.3/16 AF 12PT","RUBBER DONUT 13-14\"","SURETORQ 6PC 3/8\" NUT & BOLT EXTRACTOR SET","Sidchrome Socket 3/4\"DR 3/4\" 12PT (Imperial - Not Impact)","Sidchrome Socket 3/4\"DR 1-1/4\" 12PT (Imperial - Not Impact)","Makita 1/4\" Adjustable Locator Set w/ Bit Depth Stop for Imp","Bondhus 1/8\" Ball-End Screwdriver w/ Non-Slip Grip & Anti-Ro","CYCLONE 17X8 6/139.7 20P 110.5 SATIN BLACK ALLOY (STOCKMAN)","UNION WIRE BRUSH W/W 0.3 SW 150mm x 20W 25MB DB","ADHESIVE WEIGHT REMOVER","ABW 1-3/16 1/2 DRIVE IMPACT SOCKET","Sidchrome Socket 3/4\"DR 15/16\" 12PT (Imperial - Not Impact)","Poly Flap Disc 125x22 Coarse","Josco Brumby Fastcut Large Bar Cutting Compound (Grey)","OPEN-ENDED 12X1.5MM WHEEL NUTS 19HEX (BOX 24) 11300","EXTREME 17x8-3 6/139.7 20P BLACK 106.1CB R-FIT 1400KG ROUND","Sidchrome SPANNER RING & OPEN END 5/16\"","ABW / Sidchrome 26mm 1/2\"DR Deep Impact Socket - Metric","Sidchrome Socket 1/2\"DR 13mm 12PT DEEP (Metric - Not Impact)","JAMEC PEM FLEXIGRIP TRAY - SMALL","WHEEL WEIGHT STEEL 100G (50)","Sidchrome Socket 1/2\"DR 5/8\" 12PT (Imperial - Not Impact)","ABW 5/8 1/2 DRIVE IMPACT SOCKET","Johnson Inch/Metric Try & Miter Square Ruler - Structo-Cast","Bondhus 3/8\" L Wrench Ball End Long (BL) - ProGuard Finish I","ABW / Sidchrome 14mm 1/2\"DR Deep Impact Socket - Metric","VERMAR XR-14 RADIAL PATCHES 85X130MM (25/BOX)","MC25 Pb MOTO SPOKE WEIGHTS 25 X 25g","SURETORQ EXTRACTOR SOCKET 13mm x 3/8 Dr","Kincrome SPANNER D/R 13/16\" x 7/8\"","WHEEL WEIGHT ALLOY 45G (100)","Sidchrome Socket 3/8\"DR 22mm 12PT (Metric - Not Impact)","LFA 1.5 - 13mm Drill Chucks 1/2-20 Mount LS2 w/ Metal Sleeve","WHEEL WEIGHT ALLOY 40G (100)","WHEEL WEIGHT ALLOY 50G (50) **NEW BAGS OF 50**","SIDCHROME ADAPTOR 3/4\"F TO 1\"M SCMT19160","KC Tools 3/8\" Female x 1/2\" Male Adaptor 3/8\" Drive - Imperi","SAMSON 600x9 10PR TT SAv MB413 Set","KINCROME 11/16\" 1/2\"DR IMPACT DEEP SOCKET","TVS 10.5/80-18 IM18 14PR TL  I-1","Airco 19mm Fine Wire Staples 97 Series Electro Galvanised -","TVS 11.00-16SL (280/90-16SL) F2 12PR TL  (4 RIB) F-2","TVS 520/85R42 AR800 169D TL  TIGERTRAC R-1W","EXTREME IMITATION 16X8 6/139.7 13N BLACK R-LOCKER 110.1cb","WHEEL WEIGHT STEEL 55G (100)","Airco 12mm Staples 71 Series Electro Galvanised Chisel Pt -","KINCROME 8MM REVERSIBLE GEAR SPANNER K030031","Sidchrome Socket 1/2\"DR 13/16\" 12PT (Imperial - Not Impact)","TVS VF520/85R42 AR4005 183D/180E TL  R-1W","Sutton 14mm / 9/16\" Ultra Bi-metal Cobalt Holesaw B-056 Heav","ACORN CHROME 1/2\" 45L LONG 21HEX (Box of 24) 13940","AUTOSOL M1 CLEANSING POLISH 75ml TUBE #1910","Josco Brumby 50mm Shaft Mounted Calico Polishing Buff BCC505","TVS 400/60-22.5 FL09 18PR  TL  I-3","Sidchrome SPANNER RING & OPEN END 3/4\"","ABW / Sidchrome 1-7/16\" 1/2DR Short Impact Socket - Imperial","CARLISLE 20X10.00-10 TURFMASTE","WHEEL WEIGHT FN 45G (50) **NEW BAGS OF 50**","SAMSON 205/85R16LT GL270A 117/115L H/wy 12PR","1/2\" SOCKET RETAINING KIT (5 PACK)","Sidchrome Socket 3/8\"DR 16mm 6PT DEEP (Metric - Not Impact)","28mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024028B","ABW / Sidchrome 13mm 1/2\"DR Standard Impact Socket - Metric","W&B 3/8\" DIGITAL ANGLE TORQUE WRENCH 378338","CP AIR BLOW GUN STANDARD NOZZLE WITH SILENCER","Sidchrome Socket 1/4\"DR 6mm 6PT (Metric - Not Impact)","VWW LEAD Pb (5+10g) STICKON THIN (100)","Kincrome SPANNER D/R 3/8\" x 7/16\"","Sidchrome SPANNER RING 1\" X 1-1/8\"","Bondhus 5/8\" L Wrench Ball End Long (BL) - ProGuard Finish I","ABW / Sidchrome 1-7/16\" 1/2 DR Deep Impact Socket - Imperial","Josco Brumby 75mm Crimped Multi-Thread Cup Brush BCC65","1/2\" Dr EXTENSION 150mm DUAL ACTION 5027150B","ALPEN MULTICUT DRILL SDS+ 8.0 x 160mm (EACH)","TVS 400/15.5(400/60-15.5) IM72 18PR TL  I-3","JAMEC PEM 300mm NOZZLE BLOW GUN 07.1067","30G L/T WEIGHT (25)","Bondhus 5mm L Wrench Hex End Long (HLL) - ProGuard Finish Me","Metabo H1 M Futuro Plus Keyless Chuck R+L 0.8 - 6.5mm Single","BONJEAR WHEEL WEIGHT CUTTER","CPE43 3\" EXTENSION FOR 1/2\"","BONDHUS L Wrench - Ball End Long (BL) 1/2 (Imperial)","Sidchrome Socket 1/4\"DR 12mm 6PT (Metric - Not Impact)","EXTREME 15X7 5/114.3 0P BLACK 75CB","Bondhus 7/32\" L Wrench Ball End Long (BL) - ProGuard Finish","29mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024029B","KINCROME 13MM 1/2\"DR HEX IMPACT SOCKET","Bondhus 3/4\" L Wrench Ball End Long (BL) - ProGuard Finish I","Sidchrome Socket 1/4\"DR 10mm 6PT DEEP (Metric - Not Impact)","WHEEL 8.25x22.5 10/285 Polished Alloy Rim","PCL 30MM IMPACT SOCKET (3/4\" DRIVE)","PNEUTEK BEAD KEEPERS (PAIR)","BONDHUS L Wrench - Ball End Long (BL) 7/16 (Imperial)","SAMSON ERC600-9-10 SuperLug OB502 Solid","ABW / Sidchrome 22mm 1/2\"DR Deep Impact Socket - Metric","EXTREME IMITATION 16X8-3 6/114.3 0P BLACK D-LOCKER 66.1MM","Sidchrome Socket 3/8\"DR 17mm 6PT DEEP (Metric - Not Impact)","Bondhus 5.5mm L Wrench Hex Short Singles (HLS) - ProGuard Fi","EXTREME IMITATION 15X10 6/139.7 44N BLACK D-LOCKER 110.1","KOKEN 25mm 12Pt 1/2\" SOCKET DEEP","Sidchrome Socket Extension","ACTION 1/2\"DR x 250mm EXTENSION BAR (SB TYPE) 64521250","SMOOTHIE 15X7 10/114.3/120.65 6P CHROME 81cb","G/YEAR 225/45R19 96W EAGLE F1 ASYMMETRIC 3 (*) ROF","Sidchrome 7 x 60mm 1/2\" Drive Inhex Socket - Metric 14281","Bondhus 10mm L Wrench Ball End Long (BL) - ProGuard Finish M","Sidchrome Socket 1/2\"DR 15mm 12PT DEEP (Metric - Not Impact)","Josco Brumby 75mm Spindle-Mounted Crimped Cup Brush BCC75","KINCROME 17MM REVERSIBLE GEAR SPANNER K030040","KOKEN 29mm 12Pt 1/2\" SOCKET DEEP","Intech 3/8\" x 24 to Hex 1/4\\148 Drill Chuck Adaptor DCTA54","KOKEN 28mm 12Pt 1/2\" SOCKET DEEP","Wheel Bolt 14x1.5mm H17 S28 (28mm Thread) Chrome (20) 93728","Kincrome SPANNER COMB 1\" CARDED","Sidchrome Socket 1/4\"DR 9mm 6PT (Metric - Not Impact)","WHEEL WEIGHT TOYOTA STEEL 120g (50)","KINCROME 5/32\" 1/2\"DR HEX IMPACT SOCKET","WHEEL WEIGHT TOYOTA STEEL 100g (50)","ABW / Sidchrome 1\" 1/2\"DR Deep Impact Socket - Imperial","SNAPPY 3/8 Stop Collar","SURETORQ EXTRACTOR SOCKET 10mm x 3/8 Dr","JOHNSON LEVEL TORPEDO Magnetic 230mm","TYRE INFLATOR TDR2000","19mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522019","Kincrome SPANNER D/R 1/4\" x 5/16\"","Sidchrome Socket 1/2\"DR 1\" 12PT (Imperial - Not Impact)","Sidchrome SPANNER RING & OPEN END 12MM","SIDCHROME 3/4\"DR UNI JOINT SCMT15957","Bondhus 9/16\" L Wrench Hex End Short Singles (HLS) - ProGuar","ABW / Sidchrome 1-5/16\" 1/2 DR Deep Impact Socket - Imperial","BONDHUS L Wrench - Hex Long (HLL) 7/16 (Imperial)","ALPEN MULTICUT DRILL PM3 SET (3PC)","EXTREME 18X8 6/139.7 30P BLACK 110.1cb D-FIT","21mm SQ X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 63045021","UNION TAPERED SPINDLE SET R/H 5/8\" & 16mm","Josco Brumby 25mm Spindle-Mounted Crimped Cup Brush BCC25","EXTREME 16x8-3 6/139.7 35P BLACK 106.1CB R-FIT ROUND","3/8\" REVERSIBLE AIR DRILL DS-132","Intech 1-10mm Keyless Drill Chuck w/ Hand Tightening Plastic","3/8\" Dr EXTENSION 150mm DUAL ACTION 5017150B","Sidchrome Socket 1/2\"DR 1-1/8\" 12PT (Imperial - Not Impact)","HOSE REEL HELIX PRO SERIES EXTREME AIR 20M 58.5089","EXTREME 16x7-3 5/114.3 20P WHITE 73.1CB","Sidchrome Socket 1/2\"DR 8mm 12PT (Metric - Not Impact)","Sidchrome Socket 1/2\" DR 1-1/4\" 12PT (Imperial - Not Impact)","ABW / Sidchrome 1/2\" 1/2\"DR Deep Impact Socket - Metric","P&N Quick Change Drill 1/4 HEX 5/64 BRIGHT","STANLEY REplacement Chalk Line 100FT/30m","SURETORQ 5PC 3/8\" DEEP DRIVE NUT & BOLT EXTRACTOR SET","WIRE BRUSH WHEEL 50X19MM","DELI 570-8 (6) S378 T&T","Sidchrome INHEX SSPLINE SOCKET","Sidchrome Socket 1/4\"DR 8mm 6PT (Metric - Not Impact)","KINCROME 24mm SOCKET DEEP 1/2DR  (MP)","Sidchrome Socket 3/8\"DR 7mm 6PT (Metric - Not Impact)","Xcelite 1.5 x 102mm Allen Hex End Socket Screwdriver for Rec","ARMEG Pilot Drill 11x130mm","Airco 22mm Fine Wire Staples 97 Series Electro Galvanised -","Josco Brumby 50mm Spindle-Mounted Crimped Cup Brush BCC50","EXTREME 16x8 5/150 50N WHITE 110.1CB 1200kg","Kincrome SPANNER D/R 11/16\" x 3/4\"","ARMEG 300mm Light Weight Core Drill Adaptor","Sidchrome Socket 3/4\"DR 29mm 12PT (Metric - Not Impact)","Sidchrome Socket 1/2\"DR 3/4\" 12PT (Imperial - Not Impact)","KENDA 10.00-20(16)(TYRE & TUBE & FLAP)","Kincrome SPANNER COMB 11/16\" CARDED","ABW / Sidchrome 15/16\" 1/2\"DR Std Impact Socket - Imperial","TYRE LEVER - 300MM","KINCROME 20MM 1/2\"DR HEX IMPACT SOCKET","Sidchrome Socket 1/2\"DR 18mm 12PT DEEP (Metric - Not Impact)","KOKEN 23mm 12Pt 1/2\" SOCKET DEEP","ARMEG 65mm Heavy Duty Core Drill","Norton 305mm Speed-Grip Converter Kit No Holes - Self Adhesi","TVS 700/50-22.5 FL09 16PR TL  I-3","TVS 500/60-22.5(19.5/60-22.5)  FL09 163 A8/151 A8 16PR TL","TVS 18.4-30 TR45 10PR TL  R-1","TVS 14.9-38 TR45 8PR TL  R-1","TVS 10-16.5 ST18 10PR TL  L-4","SAMSON 14/70-20 14P C1 SET","Josco Brumby Multishine Large Bar Polishing Compound (Light","Sidchrome SPANNER RING & OPEN END 13MM","13mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024013B","UNION WIRE BRUSH B/W 0.5 T/K 125mm M14 DB","EXTREME 17x8-3 5/120 30P BLACK 65.1CB 1400KG TRIANGLE","INTECH Firm Joint Caplier - Inside 200mm","GEARWRENCH 9MM FLEX RATCHETING WRENCH","Sidchrome Socket 3/4\"DR 1-1/8\" 12PT (Imperial - Not Impact)","ABW / Sidchrome 1-3/8\" 1/2\"DR Deep Impact Socket - Imperial","UNION BUFFING WHEEL SISAL 200 1S","10PCS 6PT 1/2\" Dr 10-19 METRIC BLISTER SET - DA 10241001","WHEEL WEIGHT STEEL 80G (50)","85G L/T WEIGHT (25)","LFA 0.5-10mm Drill Chuck Keyed for Port. Power Tools w/ Meta","170G L/T WEIGHT (25)","TVS 620/75R26 HS1000 166A8/B TL TIGERTRAC  R-1","21mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014021B","26mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024026B","Sidchrome Socket 1/2\"DR 12mm 12PT DEEP (Metric - Not Impact)","SUMMIT 275/60R20 TRAIL CLIMBER H/T 115T OWL","10PCS 6PT 1/2\" Dr 10-19 DEEP METRIC BLISTER - DA 11241001","UNION TAPERED SPINDLE SET L/H 5/8\" & 16mm","Kincrome SPANNER COMB 24mm CARDED","TVS 12.5/80-18(335/80-18)BL27 12PR TL EUROGR R-4","TVS 520/85R38 AR800 169D/172A8 TL TIGERTRAC R-1W","TORO 13/6.50-6","Kin 13 MAG TOOL HOLDER\"","ACTION 85PCS 3/8\"dr 6-POINT SAE & METRIC MASTER SOCKET SET","SPLINE NUT SET CHROME 1/2\" (Box of 20 + Socket) 49034","Bondhus 5/32\" L Wrench Hex End Long (HLL) - ProGuard Finish","DURO 500-10 HF214 HWAY","WANDA 215/40-12 (4) P-825 H/Way","DELI 4.10/3.50-6 (4)","STANLEY SurForm","DIAMOND FLEX 20mm Core Drill Bit","TVS 31x13.50-15 I-09 10PR TL  I-1","TVS 600/50-22.5(23.5/50-22.5) FL 09 16PR TL  I-3","SURETORQ Adapter 1/4 Hex  X 1/4 Sq (50mm)","SURETORQ Adapter 1/4 Hex  X 3/8 Sq (50mm)","SURETORQ Extractor 5mm x 1/4\" Dr Hex Shank","KINCROME TORQUEMASTER 128PCE MASTER BIT & DRIVER","Sidchrome Socket 1/4\"DR 13mm 6PT DEEP (Metric - Not Impact)","ABW / Sidchrome 22mm 1/2\"DR Standard Impact Socket - Metric","AUTOSOL HARD WAX 500ml #3010","UNION BUFFING WHEEL STITCHED RAG 200 1S","TIANLI 29.5-25  L5 SLP 28PLY","Sidchrome Socket 3/4\"DR 19mm 12PT (Metric - Not Impact)","Bondhus 4.5mm L Wrench Hex End Long (HLL) - ProGuard Finish","SNAPPY DRILL BIT to Suit 45111","Bondhus 9mm L Wrench Hex End Long (HLL) - ProGuard Finish Me","KINCROME 11MM SINGLE WAY GEAR SPANNER K3111","4PC 1/2\" Dr Wheel Nut Deep Impact SKT Set","ALPEN MULTICUT DRILL SDS+ 5.0 x 110mm (EACH)","SURETORQ EXTRACTOR SOCKET 11/16\" x 3/8 Dr","UNION WIRE BRUSH W/W 0.3 SW 200mm x 25W 25MB DB","Intech 3/8\" x 24 Drill Chuck Adaptor Male DCTA51","TYRE GAUGE LONG STRAIGHT CHUCK","Sidchrome SPRN RING & OPEN END 3/8AF 440 SERIES","SAMSON 7.50-15 14PR SAv C-1/C-1A Set 14PR","Intech Intech Hard Hat Lanyard","Sidchrome Socket 1/2\"DR 26mm 12PT (Metric - Not Impact)","TUBE PATCH ROUND TG100MM (10)","Sidchrome Socket 1/2\"DR 17mm 12PT DEEP (Metric - Not Impact)","9mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014009B","12mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124012B","Josco Brush Cup CR 60 6 Spindle 0.30","TVS 620/55R26.5 FL918 166D TL  I-3","Sidchrome Socket 3/8\"DR 19mm 12PT (Metric - Not Impact)","SMOOTHIE 15X6 10/114.3/120.65 6P CHROME WITH BLACK TRIM 81cb","Dremel Mandrel 3.2mm Shank w/ Screw-Like Head for Dremel 414","CP 3/4\" EXT BAR 13\"","TVS 600/55-22.5 FL09 16PR TL  I-3","TG6-09 BIAS PLY 228X228 each","JOSCO LAMBSWOOL BONNET 125","TVS 540/65R28 AR600 142D/145A8 TL TIGERTRAC R-1W","Sidchrome Socket 3/4\"DR 22mm 12PT (Metric - Not Impact)","Sidchrome Socket 3/4\"DR 1\" 12PT (Imperial - Not Impact)","UNION WIRE BRUSH F/TW 0.5 T/K 125mm 22B DB","CARLISLE 13X5.00-6 SMOOTH","MARKAL \"B\" PAINTSTICK BLACK (12)","TIANLI 17.5R25**  TUL580 SMOOTH","Sidchrome Socket 1/2\"DR 23mm 12PT (Metric - Not Impact)","34mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024034B","ABW / Sidchrome 25mm 1/2\"DR Deep Impact Socket - Metric","Sidchrome Socket 1/4\"DR 1/4\" 6PT (Imperial - Not Impact)","Metabo H2 Futuro Plus Keyless Chuck R+L 10mm w/ Plastic Doub","Bordo 210mm Bi-Metal Hole Saw HSS Cobalt Cutting Edge Premiu","TVS 12-16.5 ST18 12PR TL  L-4","SURETORQ Adapter 1/4 Hex  X 1/2 Sq (50mm)","7PC 1/2\" Dr Metric Tube Deep Impact SKT Set","Paslode In-Car Power Adaptor Charger for Paslode & Spit Cord","Sidchrome Socket 3/4\"DR 27mm 12PT (Metric - Not Impact)","Sidchrome SPANNER RING & OPEN END 16MM","Metabo 10mm Futuro R Keyless Chuck 1/2\"-20 UNF for Clockwise","Sidchrome Socket 3/4\"DR 1-11/16\" 12PT","KINCROME 1/2\" 1/2\"DR HEX IMPACT SOCKET","Sidchrome Socket 1/4\"DR 13mm 6PT (Metric - Not Impact)","Sidchrome Socket 1/2\"DR 1/2\" 12PT (Imperial - Not Impact)","TVS 620/40R22.5 FL918 154D TL  I-3","Sidchrome Socket 3/8\"DR 5/8\" 6PT DEEP","TVS 10.5/80-18 (265/80-16) MT45 14PR TL  R-4","TVS 16.9-28(430/85-28) TR09 12PR TL  R-1","TVS 15.5/80-24(400/80-24) IM54 16PR TL  R-4","Sidchrome SPANNER RING & OPEN END 15MM","COMMAND FLUX & BRUSH","JAMEC PEM TDR5000 DIAL GAUGE INFLATOR","Josco 75mm Red Abrasive Nylon Cup Brush JAC75R","ABW / Sidchrome 1/2\" 1/2\"DR Standard Impact Socket - Metric","Kincrome 50 Litres Filter Cloth Bag Reusable for KP704 - 3 P","10mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014010B","10PC 1/2\" Dr 6pt Metric Deep Impact SKT Set 21mm-36mm","ACORN CHROME 1/2\" 35L SHORT 19HEX (Box of 20) 13700","22mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520022","EXTREME 16X7-3 6/139.7 42P BLACK 106.1cb","ARMEG Integral Core Adaptor to SDs MAx","ABW / Sidchrome 15/16 1/2\"DR Deep Impact Socket - Imperial","ABW 1-3/8 1/2 DRIVE IMPACT SOCKET","10PCS 6PT 1/2\" Dr 10-24 METRIC BLISTER SET - DA 10241002","ARMEG Light Weight Core Drill","27mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542027","PMT PMT VADAR Hard Hat Browguaqrd with EarMuf","EXTREME 15X6-2 5/112 30P WHITE D-HOLE 65cb","Intech 10mm x 1/4\" Hex Drill Chuck Keyless S2 Key w/ Screw D","CARLISLE 11X4.00-5 (4) SMOOTH","EXTREME 15X7 *5/139.7* 13N BLACK TRIANGLE 110.1CB","ALPEN MULTICUT DRILL SDS+ 10.0 x 160mm (EACH)","SIDCHROME 3/4\"DR EXTENSION 400mm SCMT15953","GEARWRENCH 11/16\" X 3/4\" DOUBLE BOX RATCHETING WRENCH","TVS 480/45-17 FL09 134A8 146A8  TL  I-3","Bondhus 3/4\" L Wrench Hex End Long Singles (HLS) - ProGuard","OX Ultimate 4inch Diamond Blade","12PC 3/8\" Dr 6pt Metric Deep Impact SKT Set 8mm-22mm DB","SURETORQ Power Bit SQ2 (50mm) 2 Pack","SUMMIT LT35X12.5R20 10PR MUD HOG TL 121 Q","UNION DISC SC80G 50mm RL (5PK)","17mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522017","Sidchrome SPANNER OE 20 X 22MM","Sidchrome Socket 1/2\"DR 20mm 12PT (Metric - Not Impact)","Josco Brumby 50mm Shaft Mounted Mushroom Calico Polishing Bu","UNION DISC SC320G 50mm RL (5PK)","Norton Foundry / Portable Grinding Wheel Straight 125 x 25 x","CP SILENCER EXHAUST","23mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522023","CLUTCH GREASE 14OZ CART","6PC 3/8\" Dr Triple Square Driver Set","STICK-ON Fe(5G+10G)*4 (100) BOX GREY","ALPEN MULTICUT DRILL 10.0 x 400mm (EACH)","WHEEL WEIGHT FN 20G (100)","ARMEG Integral Core Adaptor to SDS","ABW / Sidchrome 8mm 1/2\"DR Standard Impact Socket - Metric","11mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024011B","JOURNEY 13/650-6 (4) SMOOTH P607","11mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124011B","ALPEN MULTICUT DRILL SDS+ 12.0 x 210mm (EACH)","Metabo SDS-Plus Hammer Chuck for UHE 28 Plus KHE 28 Plus 631","Josco Brumby 38mm Shaft Mounted Mushroom Calico Polishing Bu","MC20 Pb MOTO SPOKE WEIGHTS 25 X 20g","JOURNEY 9/350-4 (4) SMOOTH P607 TL","KINCROME 18MM REVERSIBLE GEAR SPANNER K030041","Sidchrome Socket 3/4\"DR 32mm 12PT (Metric - Not Impact)","SMOOTHIE 15X6 10/114.3/120.65 6P GLOSS BLACK 81cb","KINCROME 16MM 1/2\"DR HEX IMPACT SOCKET","Josco Brumby 50mm Shaft Mounted Cone Calico Polishing Buff B","BKT 7.50-15 (14) PL-801 (TYRE/TUBE/FLAP)","KINCROME SOCKET DEEP 1/2DR 1-1/8 (MP)","ALPEN MULTICUT DRILL SDS+ 8.0 x 210mm (EACH)","KINCROME 15MM 1/2\"DR HEX IMPACT SOCKET","OX DRIFT Key","Sidchrome Socket 3/4\"DR 1-3/4\" 12PT (Imperial - Not Impact)","EXTREME 15x6 6/139.7 0P WHITE 111CB","SURETORQ Power Bit 6.0 (50mm) 2 Pack","Josco Brumby SSX Polishing Compound (Green)","ABW / Sidchrome 27mm 3/4\"DR Standard Impact Socket - Metric","25mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124025B","3/8\" Dr EXTENSION 75mm DUAL ACTION 5017075B","ALPEN MULTICUT DRILL 10.0 x 200mm (EACH)","KINCROME 5/16\" SINGLE WAY GEAR SPANNER K3401","DELI 600-9 (10) HWY w/ Tube","EXTREME 17x8-3 6/139.7 30P BLACK 106.1CB 1400KG TRIANGLE","WHEEL WEIGHT STEEL 60G (50) **NEW BAGS OF 50**","OX Trade 4\" Diamond Blade - General Purpose/Concrete","KINCROME SOCKET DEEP 1/2DR 7/8 (MP)","KOKEN 26mm 12Pt 1/2\" SOCKET DEEP","Bondhus 14mm L Wrench Ball End Long (BL) - ProGuard Finish M","ARMEG 100mm Heavy Duty Core Drill Adaptor","Sidchrome Socket 1/2\"DR 11mm 12PT DEEP (Metric - Not Impact)","SIDCHROME 3/4\"DR SLIDING T-HANDLE 500mm SCMT15955","ALPEN MULTICUT DRILL SDS+ 10.0 x 210mm (EACH)","ABW / Sidchrome 29mm 1/2\"DR Deep Impact Socket - Metric","ALPEN MULTICUT DRILL SDS+ 6.0 x 110mm (EACH)","UNION DISC SC120G 50mm RL (5PK)","WHEEL WEIGHT STEEL 55G (50) **NEW BAGS OF 50**","BONDHUS L Wrench - Hex Short (HLS) 5/64 (Imperial)","ABW / Sidchrome 23mm 1/2\"DR Deep Impact Socket - Metric","DUAL ACTION QUICK RELEASE RATCHET 3/8\"DR","59PCE 6PT 3/8\" Dr METRIC & IMP BOX SET - DA 10145901","DUAL ACTION 40PCE 6PT 1/2\" Dr METRIC & SAE SET 14244001","KOKEN 27mm 12Pt 1/2\" SOCKET DEEP","JOURNEY 205/50X10 (4) P823","JOURNEY 4.10/3.50-5 P605 (4) DIAMOND","KINCROME Impact Bits 20 PCE SQUARE #2","UNION WIRE BRUSH C/W 0.35 T/K 75mm M14 DB-A","ALPEN PRO DRILL KP25 (SET)","ALPEN MULTICUT DRILL 8.0 x 200mm (EACH)","DUAL ACTION 27PCE 6PT 3/8\" Dr SAE IMPACT SOCKET SET 12142701","Josco 25mm Red Abrasive Nylon Cup Brush JAC25R","40mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542040","EXTREME IMITATION 16X8-3 6/139.7 0P BLACK D-LOCKER 110.1","SURETORQ Power Bit 6.0 (150mm)","Bondhus 5/8\" L Wrench Hex End Long Singles (HLS) - ProGuard","UNION WIRE BRUSH C/W 0.5 T/K 65mm M14 DB-A","COMMAND 240V 40W SOLDERING IRON","OX SDS Adaptor","ARMEG 100mm Light Weight Core Drill Adaptor","EXTREME 14X6 5/114.3 14P WHITE 84CB (BOX) (FORD FITMENT)","GEARWRENCH 5/8\" RATCHETING WRENCH","Sidchrome SPRN RING & OPEN END 5/8AF 440 SERIES","29PCE 6PT 1/2\" Dr METRIC DEEP BOX SET - DA 11242901","KINCROME 3/8\" SINGLE WAY GEAR SPANNER K3402","33mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520033","Josco Brumby 65mm Cone Polishing Buff BCT65","KINCROME 7/16\" 1/2\"DR HEX IMPACT SOCKET","Josco Brush Cup CR 75xMT 0.30","Josco Brumby 12 Piece Multi Purpose Drill Accessory Kit BDAK","PCL BLOWGUN WITH 9\" EXTENSION","EXTREME 16x8 6/139.7 13N BLACK 110.1CB TRIANGLE","14mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024014B","25mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024025B","33mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60540033","Xcelite 6 x 102mm Allen Hex End Socket Screwdriver w/ Ballpo","Disston 60mm Remgrit Holesaw Carbide Grit Edged - 29mm Max C","ADVANCE 650-10 SOLID PREMIA","5PC  Impact Adapter Accessories Set Steel Ball Type","Bostitch 11/16\" Staples Galv C-Ring/Hog 16-Gauge Sharp Pt -","ABW / Sidchrome 1-1/4\" 1/2\"DR Deep Impact Socket - Imperial","UNION HAND BRUSH 0.35 STAINLESS STEEL WIRE BL","UNION TB-SW 0.2 13MM HEX/SK DB","TR413 - VALVE UNASSEMBLED (100)","Bondhus 1.27mm L Wrench Ball End Long (BL) - ProGuard Finish","9 PC 3/8\" Dr Impact Internal Torx Driver Set T20 - T55","Kincrome FLARE NUT SPANNER 5/8\" X 11/16\"","Promac 30mm Hole Saw TCT 12mm Deep for Sheet Metal & Stainle","BKT 700-12 (6) TL AS-504","10pc Magnetic Socket Insert Set Imperial","POWERBUILT 19MM DOUBLE BOX RATHETING WRENCH #WRT0229","Sidchrome SPANNER RING 3/4\" X 7/8\"","JOURNEY 20/10.00-10  P322","UNION HAND BRUSH SET 0.30 GREEN PVC GB","UNION FELT BUFFING SET (5PCS) DB","UNION WIRE BRUSH B/W SW SS316 100mm M14 x 2 DB-A","SMOOTHIE 15X10 10/114.3/120.65 25N GLOSS BLACK 81cb","Trend 270mm x 35mm Combination Router Base Side Fence 15mm T","UNION WIRE BRUSH W/W 0.3 SS316 150mm x 20W 25MB DB","JOURNEY 15x6.00-6 (10) P508 MULTI RIB HD","Journey 500x10 P821","MAXXIS 18X8.50-10 PROTECH M9227","TVS 600/65R28 AR2000 154D/157A8 TL  R-1W","TVS 600/70R30 AR2000 158D/161A8 TL  R-1W","18PC 3/4\" Dr 12pt IMP SAE STD IMPACT SOCKET SET 3/4\"-2\"","SURETORQ Power Bit SQ2 (50mm) (15 Tic-Tac Box)","SURETORQ Insert Bit SQ3 (2 Pack)","14mm X 3/8\" Dr Sparkplug Mag SOCKET - DUAL ACTION 5117014B","10mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124010B","DUAL ACTION 23PCE 6PT 1/4\" Dr SAE IMPACT SOCKET SET 12042301","26mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520026","8PC 1/2\" Dr 12PT Metric Deep Axle Nut Impact SKT Set","W&B 3/8\" DEFLECTING BEAM TORQUE WRENCH 321500","EXTREME IMITATION 17X9 5/127 0P BLACK D-LOCKER 71.5","Sidchrome Socket 1/2\"DR 31mm 12PT (Metric - Not Impact)","Sidchrome Socket 1/2\"DR 9/16\" 12PT (Imperial - Not Impact)","UNION HAND BRUSH 0.5 GREEN PVC YW","UNION TB-NW 25MM","UNION DISC S/ABRASIVE 75mm RL","UNION PAD QUICK-LOC 50 X 6mm SHAFT","UNION DISC FELT 75mm RL","UNION WIRE BRUSH E/B SW 0.3 SS316 20mm RD/SK DB","Bondhus 2.5mm L Wrench Hex End Long (HLL) - ProGuard Finish","Proxxon Micromot Tool Holder Clamp for Stationary Use NO2841","Sidchrome 5/16\" x 60mm 1/2\" Drive Inhex Socket - Imperial 14","JOURNEY 13/500-6 SLIDE","MAXXIS 12.00-20 (20) TOUGHGUARD (TYRE,TUBE,BAND)","JOURNEY 26X9.00-12 (12) WARLOCK P350","UNION WIRE BRUSH C/W 0.5 T/K 100mm M14 DB","UNI PATCH SQUARE TGU55MM (40)","Sidchrome Socket 3/4\"DR 1-13/16\" 12PT","17mm x 3/8 Dr 6pt DEEP WALL SOCKET - DUAL ACTION 1114017B","20mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124020B","20mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014020B","UNION WIRE BRUSH B/W SW SS316 125mm M14 x 2 DB","29mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124029B","16mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522016","29PCE 6PT 1/2\" Dr METRIC STD BOX SET - DA 10242901","DUAL ACTION 22PCS 12PT 1/2\" Dr SAE DEEP SOCKET SET 17242201","UNION DISC SC120G 75mm RL (5PK)","24mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522024","14PC 1/2\" Dr 12PT Metric STD Impact SKT Set 10mm - 27mm","JAMEC PEM DUAL CHUCK NON SEALING 03.0465","W&B 1/4\" MICROMETER TORQUE WRENCH 370000B","SURETORQ Insert Bit Slot 8-10 (2 Pack)","ABW / Sidchrome 26mm 1/2\"DR Standard Impact Socket - Metric","MALCO Punch & Die for Stud Crimper","MAKITA Impact Square #2 x 25mm Insert Bit 2PK","UNION F/W SUPER ABRASIVE FIBRE 100mm X 16B DB","Josco Brush Wheel CR 150x19xMB 0.35","ARMOR 7-14 R1","19mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60720019","ABW / Sidchrome 3/4\"  3/4\"DR Standard Impact Socket - Imperi","KINCROME 4MM 1/2\"DR HEX IMPACT SOCKET","Bondhus 1.5mm L Wrench Ball End Long (BL) - ProGuard Finish","Bordo 177mm Bi-Metal Hole Saw HSS Cobalt Cutting Edge Premiu","UNION WIRE BRUSH B/W 0.5 T/K SS316 125mm M14 DB","BKT 11.00-16 (12) F3","UNION WIRE BRUSH F/TW 0.5 T/K SS316 125mm 22B DB","KENDA 25X1250-10 BEARCLAW","ARMEG Core Drill 45mm","DURO 18X8.50-8 6 TL HF224","TVS 23.1-26 TM09 168A8 16PR HD TL  R-3","WHEEL WEIGHT FN 55G (50) **NEW BAGS OF 50**","NOVUM 28x9-15 SOLID","SMOOTHIE 15X6 10/114.3/120.65 6P CHROME 81cb","EXTREME 17x8-3 6/114.3 35P BLACK 66.1CB S8-FIT 1400KG SOFT8","TIANLI 18.4-30  SEWP(ST) 16 TT","WANLI 185R15C 8PR S-2010(DISCONTINUED)","WANLI 215/75R17.5 16PR SDR01 135/133J","Sidchrome Socket 1/2\"DR 29mm 12PT (Metric - Not Impact)","30mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520030","TIANLI 29.5-29  DNR(S) 28PLY","ACTION 3/4\"DR x 250mm EXTENSION BAR (SB TYPE) 64541250","W&B 1/2\" SCREEN MICROMETER TORQUE WRENCH 334451","7PC 3/8\" Dr Metric Impact Hex Driver Set 4mm - 12mm","Sidchrome SPANNER RING & OPEN END 9MM","GTS DRILL BIT 8MM","12mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024012B","9mm x 3/8 Dr 6pt DEEP WALL SOCKET - DUAL ACTION 1114009B","23mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024023B","HOSE ASSEMBLY FOR 11.0723/11.0734","UNION TB-SW 32MM","12mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60720012","1/2\"F x 3/4\"M ADAPTOR DUAL ACTION 4924048B","UNION W/B ABRASIVE FIBRE 100mm HEX/SX DB","KOKEN 20mm 12Pt 1/2\" SOCKET DEEP","Sidchrome Socket 3/8\"DR 12mm 6PT DEEP (Metric - Not Impact)","DUAL ACTION 23PCE 12PT 1/4\" Dr SAE IMPACT SET 16042301","Sidchrome SPANNER RING 13/16\" X 7/8\"","UNION WIRE BRUSH W/W 0.3 SS316 200mm x 25W 25MB DB","UNION HAND BRUSH SET 0.2 STEEL WIRE GB","35mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520035","UNION TB-SW 25MM","Promac 45mm Hole Saw TCT 12mm Deep for Sheet Metal & Stainle","OX Hex Adaptor","Sidchrome SPANNER RING 5/8\" X 11/16\"","ARMEG 50mm Heavy Duty Core Drill","UNION TB-SW 13MM","ASCENSO 11.2-24 8ply R1 116A8","UNION TB-SW 0.2 16MM HEX/SK DB","TIANLI 650/65R42  AG-R 165D168A","UNION COMPOUND FASTCUT - GREY","UNION TB-NW 16MM","1/4\" Dr Universal Joint DUAL ACTION 4800043B","KINCROME 3/4\" SINGLE WAY GEAR SPANNER K3408","DELI 400-8 (6) S234 Hwy","28mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60140028","14PC 1/2\" Dr 6pt Metric Deep Impact SKT Set 10mm - 27mm","UNION WIRE BRUSH W/B SW 0.3 25mm x 8W RD/SK DB","41mm X 21mm SQ x 3/4\" Dr SOCKET - ACTION 61540041","Sidchrome SPANNER RING & OPEN END 7/16\"","19PC 1/2\" Dr 6pt SAE Deep Impact SKT Set 3/8\" - 1-1/2\"","BKT 12-16.5 (12) TL SKID POWER HD","Kincrome Tamperproof Torx Socket Set - Long Series Bit Sets","10PC 1/2\" Dr 6pt Metric Deep Impact SKT Set 10mm-19mm","UNION HAND BRUSH SET 0.3 GREEN PVC DB","TVS 14L-16.1SL I-09 12PR TL  I-1","UNION W/B ABRASIVE FIBRE 75mm HEX/SX DB","UNION WIRE BRUSH C/W SS316 75mm M14 DB-A","UNION BUFFING WHEEL CALICO 200 X 50 FOLD","UNION PAD QUICK-LOC 75 X 6mm SHAFT","BONDHUS T-Handle Hex 9/64 (Imperial)","Metabo S2 Futuro Plus Keyless Chuck R+L 13mm w/ Plastic Doub","10PC 3/8\" Dr Metric Uni STD Impact SKT Set 10mm - 19mm","Unifit 10 x Gripwell Dust Bag w/ Filter Gesadie Hoover UNI64","BKT 11.5/80-15.3 MP-567 TL","KENDA 9/3.50-4 K372","SURETORQ Insert Bit PH 1 (2 Pack)","WESTLAKE 500-10 CR823","JOURNEY 18X8.50-10 P332","UNION WIRE BRUSH W/B SW 0.3 38mm x 8W RD/SK DB","7mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014007B","21mm X 3/8\" Dr Sparkplug Mag SOCKET - DUAL ACTION 5117021B","13mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124013B","UNION WIRE BRUSH W/B SW 0.3 50mm x 8W RD/SK DB","EXTREME 14X6 5/114.3 0P WHITE 84CB (FORD FITMENT)","EXTREME 16x8 5/150 50N BLACK 110.1CB D-FIT D-HOLE 1400kg","ACORN CHROME 14X1.5mm 45L LONG 21HEX (Box of 20) 13940","KINCROME 18MM SINGLE WAY GEAR SPANNER K3118","Kincrome 11 Piece Torquemaster General Purpose Bit & Driver","11mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522011","MC15 Pb MOTO SPOKE WEIGHTS 25 X 15g","1/4\"F x 3/8\"M ADAPTOR DUAL ACTION 4901026B","UNION TB-SW 0.2 19MM HEX/SK DB","UNION TB-NW 10MM","UNION DISC FELT 50mm RL","UNION WIRE BRUSH B/W 0.5 SS316 T/K 100mm M14 DB-A","Makita 5?1/2\" Drill Bit Extension Bar w/ High Strength Steel","BONDHUS L Wrench - Hex Long (HLL) 1/2 (Imperial)","33mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542033","KINCROME 11MM REVERSIBLE GEAR SPANNER K030034","Makita 7\" Sponge Pad Hook & Loop for Polishing Applications","10PC 1/2\" Dr Metric Impact Hex Driver Set 6mm - 19mm","13PC 3/8\" Dr 6pt Metric Deep Mag Impact SKT Set 10mm - 22mm","11mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60720011","DUAL ACTION 27PCE 12PT 3/8\" Dr SAE IMPACT SET 16142701","BONDHUS L Wrench - Ball End Long (BL) 1/16 (Imperial)","13mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522013","25mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522025","12PC 3/8\" Dr IMP SAE Universal DP Impact SKT Set 5/16\" - 1\"","KINCROME 11/16\" SINGLE WAY GEAR SPANNER K3407","Kincrome Torx Socket Set - Long Series Bit Sets","UNION TB-SW 0.2 25MM HEX/SK DB","14PC 1/2\" Dr 6pt SAE STD Impact SKT Set 7/16\" - 1-1/4\"","SURETORQ Insert Bit PH 3 (2 Pack)","33mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124033B","12PC 1/2\" Dr 6pt IMP / SAE Impact SKT Set 5/16\"-1\"","6PCS 3/8\" Dr Impact Adapter & Extension Accessories Set 6401","UNION WIRE BRUSH W/B SW 0.3 19mm x 5W RD/SK DB","Kincrome SPANNER COMB 9mm CARDED","UNION WIRE BRUSH B/W 0.5 T/K 100mm M14 DB-A","Promac 75mm Hole Saw TCT 12mm Deep for Sheet Metal & Stainle","PCL 12MM DEEP IMPACT SOCKET (1/2\" DRIVE)","Sutton 19mm Multi-Purpose Holesaw TCT H1110190","Sidchrome Socket 3/4\"DR 1-3/8\" 12PT (Imperial - Not Impact)","W&B 1/2\" MICROMETER TORQUE WRENCH 373000","ABW / Sidchrome 24mm 1/2\"DR Standard Impact Socket - Metric","UNION WIRE BRUSH F/W 0.5 T/K 100mm 16B DB","SNAPPY 1/16 Drill Bit Adaptor","MAKITA DIE","TVS 16.9-30 TR45 12PR TL  R-1","BKT 10.5R20 TL MP567","20mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522020","5PCS Impact Adapter Set","Kincrome SPANNER D/R 9/16\" x 5/8\"","Kincrome BLADE SKT SET 11PCE MET","10mm X 3/8\" Dr STANDARD IMPACT SOCKET - ACTION 61110010","TUBE PATCH OVAL TG100X50 (30)","ABW 11/16 1/2 DRIVE IMPACT SOCKET","21mm x 3/8 Dr 6pt DEEP WALL SOCKET - DUAL ACTION 1114021B","DUAL ACTION 22PCS 6PT 1/2\" Dr SAE SOCKET SET 12242201","17mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60720017","UNION WIRE BRUSH E/B SW 0.3 20mm RD/SK DB","UNION W/B ABRASIVE FIBRE 50mm HEX/SX DB","28PC 1/2\" Dr 6pt SAE STD & Deep Impact SKT Set 7/16\"-1-1/4\"","UNION WIRE BRUSH E/B SW 0.3 25mm RD/SK DB","15mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522015","10PC 3/4\" Dr IMP / SAE STD Impact SKT Set 1\" - 1 5/8\"","UNION TB-SW 19MM","18PC 1/2\" Dr Metric & IMP SAE Impact Hex Driver Set","UNION TB-NW 13MM","6PC 1\" Dr Metric Impact Hex Driver Set 14mm - 27mm","SURETORQ Power Bit T15 (50mm) 2 Pack","35mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024035B","UNION TB-NW 19MM","SURETORQ Insert Bit SQ2 (15 Tic-Tac Box)","UNION W/B ABRASIVE FIBRE 65mm HEX/SX DB","UNION F/W SUPER ABRASIVE FIBRE 115mm X 22B DB","UNION TB-SW 0.2 10MM HEX/SK DB","UNION DISC SC320G 75mm RL (5PK)","29mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520029","EXTREME 17x8-3 6/114.3 35P BLACK 66.1CB R-FIT 1400KG ROUND","SIDCHROME 3/4\"DR EXTENSION 100mm SCMT15952","UNION BUFFING WHEEL CALICO 200 X 100 FOLD","UNION BUFFING WHEEL STITCHED RAG 150 1S","KINCROME 19MM REVERSIBLE GEAR SPANNER K030042","UNION HAND BRUSH SET 0.20 STEEL WIRE DB","1/2\"F x 3/8\"M ADAPTOR DUAL ACTION 4921038B","UNION DISC SC80G 75mm RL (5PK)","UNION COMPOUND SSX-GREEN","UNION TB-NW 32MM","28mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124028B","DUAL ACTION QUICK RELEASE RATCHET 1/2\"DR 250mm","Airco 22mm Brad Nail C Series Electro Galvanise Plain Shank-","Metabo 1-10mm Futuro Top Keyless Chuck R 1/2\"-20 UNF Impact","ALLIANCE 500/60-22.5 16PR TL 328 163A8 I-3 FLOTATION","12mm X 3/8\" Dr STANDARD IMPACT SOCKET - ACTION 61110012","EXTREME 17x8-3 6/139.7 20P BLACK 106.1CB D-FIT 1400KG DHOLE","13mm X 3/8\" Dr DEEP IMPACT SOCKET - ACTION 60512013","Bondhus 4.5mm L Wrench Ball End Long (BL) - ProGuard Finish","26mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522026","TIANLI 23.1-26  SL 16PR LS-2 TT","CARLISE 29x12.50-15 (6) TL TRU POWER (320/55-15)","20mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024020B","33mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522033","10PC 3/4\" Dr 6pt IMP / SAE Deep Impact SKT Set 1\" - 1-5/8\"","UNION WIRE BRUSH C/W 0.5 T/K 75mm M14 DB-A","Josco 50mm Red Coarse Abrasive Nylon Wheel Brush JAW50R","JAMEC PEM FLEXIGRIP BELT - LARGE 30\"","ABW / Sidchrome 1-1/2\" 1/2\"DR Deep Impact Socket - Imperial","EXTREME 15x6 6/139.7 0P BLACK 111CB","24mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124024B","11mm X 3/8\" Dr STANDARD IMPACT SOCKET - ACTION 61110011","10mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60720010","ACORN BLUE 1/2\" 35L SHORT 19HEX (Box of 20) 13700","KINCROME 19MM SINGLE WAY GEAR SPANNER K3119","27mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522027","TVS VF600/70R28 AR4005 173D TL  R-1W","SURETORQ 21PC SCREWDRIVER SET WITH CARRY BAG","12PC 1/2\" Dr 6pt Metric Deep Impact SKT Set 10mm-24mm","DUAL ACTION Repair Kit 3/8\" Dr Ratchet Wrench","SURETORQ Insert Bit SQ1 (2 Pack)","Ajax Mini Oiler 1/4\" NPT Air Inlet & 1/4\" NPT Outlet w/ Bras","SURETORQ Power Bit T20 (150mm)","UNION WIRE BRUSH F/W 0.4 T/K 100mm RD/SK DB","ABW / Sidchrome 9/16\" 1/2\"DR Deep Impact Socket - Imperial","EXTREME 16x8-3 6/139.7 35P BLACK 106.1CB (TRIANGLE)","W&B 14x18mm INTERCHANGEABLE SCREEN TORQUE WRENCH 334053","JOURNEY 3.00-4 (4PR) TT DIAMOND","Josco Brush Cup TK18 75x1RxMT 0.35","EXTREME 17x8-3 6/139.7 23N BLACK 110.1CB R-FIT 1400KG ROUND","UNION DISC S/ABRASIVE 50mm RL","SURETORQ Bit Holder-Mag with Clip (57mm)","UNION BUFFING WHEEL SISAL 150 1S","Sidchrome Socket 1/2\"DR 12mm 12PT (Metric - Not Impact)","Sidchrome Socket 1/2\"DR 10mm 12PT DEEP (Metric - Not Impact)","14mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124014B","Vacmaster 38L Tank Liner Bags for M-Class Dust Extractors -","ABW / Sidchrome 24mm 1/2\"DR Deep Impact Socket - Metric","DUAL ACTION 31PCE 12PT 1/4\" Dr METRIC IMPACT SET 14043101","DUAL ACTION 29PCE 12PT 1/2\" Dr METRIC DEEP  SET 15242901","25mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520025","13PC 1/2\" Dr Metric Universal STD Impact SKT Set","UNION WIRE BRUSH E/B SW 0.3 12mm RD/SK DB","7PC 3/8\" Dr IMP SAE Impact Hex Driver Set 3/16\" - 1/2\"\"","UNION F/W SUPER ABRASIVE FIBRE 125mm X 22B DB","UNION WIRE BRUSH F/W 0.5 T/K 120mm 22B DB","UNION WIRE BRUSH PE/B SW 0.3 20mm RD/SK DB","UNION WIRE BRUSH B/W SW 100mm M14 x 2 DB-A","SMOOTHIE 15X8 10/114.3/120.65 12N GLOSS BLACK 81cb","Sidchrome Socket 3/8\"DR 14mm 6PT DEEP (Metric - Not Impact)","SMOOTHIE 15X8 10/114.3/120.65 25N GLOSS BLACK 81cb","GALAXY 12.5L-15SL 12PR TL I-1 RIB IMP-544166","Sidchrome Socket 1/2\"DR 15/16\" 12PT (Imperial - Not Impact)","34mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520034","SURETORQ Adapter Set 3PCS 1/4 Hex (50mm)","ABW / Sidchrome 3/8\" 1/2\"DR Deep Impact Socket - Imperial","UNION TB-SW 10MM","UNION TB-SW 16MM","KINCROME 9MM SINGLE WAY GEAR SPANNER K3109","Bondhus 7mm L Wrench Hex Short Singles (HLS) - ProGuard Fini","30PC 1/2\" Dr Metric STD & Deep Impact SKT Set 10mm - 24mm","5PC 3/4\" Dr Impact Adapter & Extension Accessories Set","ADVANCE 825x15 (14) OB502","DEESTONE 4.00-8 D401","TVS VF710/60R30 AR4005 171D/168E TL R-1W","ABW / Sidchrome 13mm 1/2\"DR Deep Impact Socket - Metric","Josco 50mm Red Abrasive Nylon Cup Brush JAC50R","JOSCO JCC63 CRIMPED CUP BRUSH 63MM","10mm X 3/8\" Dr DEEP IMPACT SOCKET - ACTION 60512010","30mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542030","36mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60540036","SMOOTHIE 15X8 10/114.3/120.65 12N CHROME 81cb","26PC 1/2\" Dr 12PT Metric Deep Impact SKT Set 10mm - 36mm","Bondhus 2mm L Wrench Hex Short Singles (HLS) - ProGuard Fini","SURETORQ Nutsetter Magnetic 1/4\" (65mm)","SURETORQ EXTRACTOR SOCKET 12mm x 3/8 Dr","Josco Brumby Jumbo Metal Polishing Kit JMPKIT5","UNION HAND BRUSH SET 0.13 STAINLESS STEEL DB","UNION TB-SW 0.2 32MM HEX/SK DB","UNION WIRE BRUSH C/W 0.5 SS316 T/K 75mm M14 DB-A","Bondhus 11mm L Wrench Hex Short Singles (HLS) - ProGuard Fin","SURETORQ Power Bit T10 (150mm)","ABW / Sidchrome 1-7/16\" x 3/4\"DR Deep Impact Socket - IMPERI","UNION WIRE BRUSH W/B SS316 0.3 75mm x 13W RD/SK DB","ALPEN SPRINT MASTER DRILL 3.2 x 65mm (EACH)","KINCROME 9MM REVERSIBLE GEAR SPANNER K030032","ALPEN MULTICUT DRILL 12.0 x 400mm (EACH)","KINCROME SOCKET DEEP 1/2DR 1-1/16 (MP)","TVS VF720/75R42 AR4005 184D TL R-1W","12mm X 3/8\" Dr DEEP IMPACT SOCKET - ACTION 60512012","12mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522012","12PC 1/2\" Dr 6pt IMP SAE Deep Impact SKT Set 5/16\"-1\"","SIDCHROME 3/4\"DR ADJUSTABLE OFFSET HANDLE 500mm SCMT15954","11PC 1/2\" Dr Impact Internal Torx Driver Set T20 - T70","UNION WIRE BRUSH F/W 0.4 T/K 75mm RD/SK DB","double ring spanner 11 &14","8PC 1\" Dr 6Pt Metric Deep Impact SKT Set 27mm - 41mm","UNION COMPOUND MULTISHINE - BLUE","Paslode Impluse Framing Panel Nose Eliminates Wood Markings","INTECH Caplier _Spring Type - Inside","OX Professional 6\" Turbo Diamond Blade","Dremel Flange Plate Replacement - 2610021327","BKT 500-12 (6) TR-171 (R-1W+)","EXMILE 14-17.5 ESK308 LOM","Metabo Powergrip Charging Adapter Suitable for Battery Model","ARMEG Light Weight Core Drill Tommy Bar","JOURNEY 8X3.00-4 (4) P607 SMOOTH","6\" ANVIL FOR CP7780-6","UNION WIRE BRUSH C/W 0.5 T/K 125mm M14 DB","15mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014015B","UNION WIRE BRUSH W/B SW 0.3 75mm x 13W RD/SK DB","15mm x 3/8 Dr 6pt DEEP WALL SOCKET - DUAL ACTION 1114015B","14mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60720014","14PC 1/2\" Dr 6pt Metric STD Impact SKT Set 10mm - 27mm","SURETORQ Insert Bit T15 (2 Pack)","JAMEC PEM FLEXIGRIP TRAY - LARGE","W&B 3PC 100T SWIVEL RATCHET SET 163351","Kincrome SPANNER COMB 3/4\" CARDED","Sutton HOLESAW SET TCT DOWNLIGHT 9 PCE","16mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024016B","Makita Guard Roller Wheel 20 / 5903 - 415438-0","23mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60520023","35mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522035","27mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60140027","46mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60540046","10PC 1/2\" Dr 6pt IMP / SAE Deep Impact SKT Set 7/16\" - 1\"","10PC 3/4\" Dr 6pt Metric STD Impact SKT Set 21mm - 46mm","8PC 3/8\" Dr Metric Impact Hex Driver Set 3mm-12mm","JOURNEY 18X8.50-8 P332","JOURNEY 25X10.00-12 (6) P341 FARM WIZARD","MC10 Pb MOTO SPOKE WEIGHTS 25 X 10g","21mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542021","SUMMIT LT37X12.5R20 10PR MUD HOG TL 126 Q","4PC 3/4\" Metric Impact Ball End Hex Driver Set 17mm - 24mm","16mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124016B","EXTREME 17x8-3 5/120 30P BLACK 65.1CB D-FIT 1400KG DHOLE","34mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522034","W&B 3/8\" MICROMETER TORQUE WRENCH 370000","EXTREME 17x8-3 6/139.7 20P BLACK 106.1CB 1400KG TRIANGLE","Kincrome 8 Piece Bit socket -Hex 1/2\" Drive Metric K2135","LIFTEX 6.50-10/5.0 PM STD PEAKMASTER","MAXXIS 18X8.50-8 4PR M7515 POWER LUG","Weller Smoothing Pen Replacement Part 6120","AUTOSOL MOTORBIKE POLISH & WAX 50g #1080","ADVANCE 5.70-12 (8) L3","BKT 15.5-25 12PLY T/L  GR288 LOM","DELI 18/650-8","SAILUN 175/65R14C COMMERCIO VX1 (CLEARANCE)","POWERBUILT 18MM RATCHET RING SPANNER","Sidchrome Socket 1/2\"DR 27mm 12PT (Metric - Not Impact)","Sidchrome Socket 3/4\"DR 1-1/2\" 12PT (Imperial - Not Impact)","14mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014014B","15mm X 1/2\" Dr STANDARD IMPACT SOCKET - ACTION 60720015","18mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522018","3/8F x 1/4M ADAPTOR DUAL ACTION 4910032B","14PC 1\" Dr 6pt IMP SAE Deep Impact SKT Set 1-1/4\"\" - 2-1/2\"","14PC 1\" Dr 6pt Metric Deep Impact SKT Set 27mm-65mm","W&B 1/4\" DEFLECTING BEAM TORQUE WRENCH 320510","Kincrome 38 Piece Automotive Bit Socket Set K2138","35mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124035B","Intech 3/8\\148 x 24 Male to 1/2\\148 x 20 Male Drill Chuck Adaptor","Intech Drill Chuck Key KK Jacobs Part 73mm 10 Teeth DKEYKK","Lion Octopus Strap 450mm Cord w/ 5mm Latex Core & Heavy Duty","Promac 85mm Hole Saw TCT 12mm Deep for Sheet Metal & Stainle","Lion Octopus Strap 48\" Cord w/ 5mm Latex Core & Heavy Duty S","MAKITA Dust + Nozzle For Planer","24mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60140024","JOURNEY 27 - 8.50 X 15 TURF","BKT 5.00-15 (4) TF-9090","Makita Roller Pin 6 / 5903 - 415437-2","46mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542046","P&N DRILL ADAPTOR 1/4 (6.1-6.4mm)","36mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542036","CHENG SHIN 13/6.50-6","ACTION 1/2\"F X 3/4\"M ADAPTOR (SB TYPE) 64021024","JOURNEY 23X9.50-12 (4) P332 TURF","18PC 3/4\" Dr Metric STD Impact SKT Set 19mm - 50mm","14mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522014","28mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522028","26mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60140026","38mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542038","13PC 1/2\" Dr Quick Release Chuck / Hex Driver Bit Set","35mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60540035","8PC 1/2\" Dr Impact Internal Torx Driver Set T30 - T70","SURETORQ Power Bit SQ3 (50mm) 2 Pack","SURETORQ Extension 1/4 Hex x 1/4 QC Sq (65)","Bondhus 3.5mm L Wrench Ball End Long (BL) - ProGuard Finish","KINGS 600X9 (6) K364 TYRE&TUBE","EXTREME 15X7 *5/139.7* 10P BLACK TRIANGLE 110.1CB","10mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024010B","SURETORQ Insert Bit T25 (2 Pack)","DUAL ACTION 53PCE 12PT 1/4\"Dr Metric & SAE SET 14045302","Kincrome SPANNER COMB 13/16\" CARDED","Sidchrome Socket 1/2\"DR 15mm 12PT (Metric - Not Impact)","Sidchrome Socket 3/4\"DR 41mm 12PT (Metric - Not Impact)","19mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60140019","Sidchrome SPRN RING & OPEN END 1-1/8AF 440 SERIES","50mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60540050","SNAPPY Quick Chuck with SDS Shank","28PC 1/2\" Dr 6pt Metric & SAE Deep Impact SKT Set","6PC 1\" Dr IMP / SAE Impact Hex Driver Set 3/4\" - 1-3/8\"","UNION HAND BRUSH 0.35 BRASS WIRE GN","Bordo 14mm Bi-Metal Hole Saw Cobalt HSS Premium Quality 7010","INTECH Firm Joint Caplier - Inside 150mm","Stanley Line Level Flat Base Lightweight 3\" ABS Material w/","MILWAUKEE Shockwave 3/8 Square x 1/4","W&B 3/4\" DEFLECTING BEAM TORQUE WRENCH 325510","BKT 18x7-8 (16) Power Trax HD","GLOBESTAR 700-12/5.0 WT GREY NM STD","UNION PAD QUICK-LOC 75 X M10 X 1.5","ACORN RED 12X1.25mm 35L 19HEX (Box of 20) 13700","15mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024015B","KOKEN SOCKET DEEP 1/2DR X 1.1/16 AF 12PT","Sidchrome Socket 1/4\"DR 7mm 6PT (Metric - Not Impact)","Sidchrome Socket 1/2\"DR 30mm 12PT (Metric - Not Impact)","SIDCHROME 3/4\"DR EXTENSION 200mm SCMT15952","8mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014008B","16mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014016B","15mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124015B","MAKITA UNIVERSAL TCT DRILL BIT 6.5x150mm","16mm X 3/8\" Dr STANDARD IMPACT SOCKET - ACTION 61110016","38mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60540038","40mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60540040","KINCROME 3/4MM REVERSIBLE GEAR SPANNER K030018","SURETORQ Nutsetter Magnetic 5/16\" (65mm)","75X30MM KNOT WHEEL BRUSH","W&B MAGNETIZER & DEMAGNETIZER TOOL 163104","KENDA 18x9.50-8 (2) K290 Scorpion","Sidchrome Socket 1/2\"DR 1-3/16\" 12PT (Imperial - Not Impact)","Metabo 16mm Geared Chuck w/ Key & Fem Thread for Clockwise R","Metabo B12 Futuro Top Keyless Chuck R 10mm Quick Release & O","Promac 100mm Hole Saw TCT 12mm Deep for Sheet Metal & Stainl","26mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542026","Sutton 33mm / 1.5/16\" Ultra Bi-metal Cobalt Holesaw B-131 He","41mm X 3/4\" Dr DEEP IMPACT SOCKET - ACTION 60542041","18mm x 3/8 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1014018B","MAKITA UNIVERSAL TCT DRILL BIT 7x100mm","42mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60540042","MAKITA Vise Holder Assembly/GD0800C/GD0","10PC 1/2\" Dr 6pt Metric STD Impact SKT Set 10mm - 24mm","NANKANG 22X11.00-8 TURF","Sidchrome Socket 3/4\"DR 36mm 12PT (Metric - Not Impact)","BKT 570X8 8PLY ST-180 HIGHWAY","OPEN EYE NEEDLE 4\" (5 PACK)","SURETORQ Power Bit T25 (50mm) 2 Pack","ABW / Sidchrome 11mm 1/2\"DR Deep Impact Socket - Metric","SURETORQ Power Bit Slot 8-10 (50mm) 2 Pack","12PC 1/2\" DR RIBE DRIVER SET","Promac 51mm Hole Saw TCT 12mm Deep for Sheet Metal & Stainle","Airco 45mm Staples Heavy Wire BCS1500 Series Electro Galv -","Josco Brush Wheel TK30 200x1RxMB 0.50","ABW ADAPTOR 1\"DR TO 3/4\"DR","27mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024027B","Josco Brush Cup CR 75xMT 0.30 Brass","DELI 450-10 (6) S252","SUMMIT AT LT265/75R16 10PR 123/120S TRAIL CLIMBER OWL","CP7748 1/2\" IMPACT WRENCH G-SERIES","CP 3/4\"X10\" IM EXTENSION BAR","ALPEN MULTICUT DRILL 5.0 x 85mm (EACH)","34mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124034B","W&B 3/4\" LH/RH DEFLECTING BEAM TORQUE WRENCH 325520","1/2\" Dr EXTENSION 75mm DUAL ACTION 5027075B","29mm X 1/2\" Dr DEEP IMPACT SOCKET - ACTION 60522029","21mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60140021","22mm X 3/4\" Dr STANDARD IMPACT SOCKET - ACTION 60140022","10PC 1/2\" Dr 6pt IMP / SAE STD Impact SKT Set 7/16\" - 1\"","SURETORQ Power Bit PH3 (50mm) 2 Pack","UNION WIRE BRUSH C/W SW 75mm M14 DB-A","Sidchrome Socket 1/2\"DR 28mm 12PT (Metric - Not Impact)","Sidchrome Socket 3/4\"DR 50mm 12PT (Metric - Not Impact)","CP7748 1/2\" IMPACT WRENCH G-SERIES + PROMO BAG","JOURNEY 23/8.50X12 6PLY P332","CARLISLE 23x8.50-12(4) TRU POWER","Bondhus 1.27mm L Wrench Hex End Long (HLL) - ProGuard Finish","9PC 3/8\" Dr Thin Wall External Torx Impact SKT Set E5 - E16","ALPEN SPRINT MASTER DRILL KP25 (SET)","Kincrome SPANNER COMB 13MM CARDED","Airco 56mm Brad Nail C Series Electro Galvanise Plain Shank-","Lion 6pcs Mini Snap Off Blades for LA144B2 LT144B7","Proxxon Micromot Polishing Felt Medium Hard 50mm for Polishi","MAKITA Impact Square #1 x 25mm Insert Bit 2PK","11mm x 3/8 Dr 6pt DEEP WALL SOCKET - DUAL ACTION 1114011B","KINCROME 5/8\" SINGLE WAY GEAR SPANNER K3406","Bondhus 14mm L Wrench Hex End Long (HLL) - ProGuard Finish M","OX 12mm Guide Rod","CHENG SHIN 260x85","19mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024019B","JOURNEY 11X4.00-5(4) P508 MULTIRIB","27mm x 1/2 Dr 6pt THIN DEEP SOCKET - DUAL ACTION 1124027B","33mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024033B","36mm x 1/2 Dr 6pt THIN WALL SOCKET - DUAL ACTION 1024036B","Sidchrome Socket 3/8\"DR 19mm 6PT DEEP (Metric - Not Impact)","10PC 1/2\" Dr 6pt Metric STD Impact SKT Set 10mm - 19mm","16PC 1/2\" Dr Thin External / Internal Torx Impact SKT Set","Metabo B18 Geared Chuck 16mm w/ Key & Fem Thread for Clockwi","8PC Impact Adaptor Set Steel Ball Retainer Type","Sidchrome SPANNER RING & OPEN END 1-3/8\"","ARMEG 80mm Heavy Duty Core Drill","APLUS 215/70R14 96H A606","BKT 380/70R24 125/A8/B TL RT-765 AGRI MAX","SUMMIT AT LT225/75R16 10PR 115/112S TRAIL CLIMBER BSW","Sidchrome SPANNER RING & OPEN END 1\"","ZETA 255/55R19","ZESTINO 235/45R17"];

// Chapel Corner: use the real, human-curated exclusion list (supplied 2026-07-15,
// "chapel corner - other.txt") instead of regex heuristics. Validated against
// BigQuery: reproduces the ORIGINAL static dashboard's Chapel Corner numbers
// exactly (tyre: matched 1136 / unmatched 1520; other: 1773 unmatched, 0 matched)
// — this list IS the "user-identified exclusion list" the original dashboard
// footnote referenced. Only two buckets for this supplier: other (in the list,
// includes wheels/valves/tools/niche tyres per the human classification) and
// tyre (everything else). No separate wheel bucket here.
function stockProductTypesSQLExactMatch(names) {
  return 'SELECT\n' +
    '  CASE WHEN supplier_description IN (' + sqlQuoteList(CHAPELCORNER_EXCLUSION_LIST) + ')\n' +
    "    THEN 'other' ELSE 'tyre' END AS product_type,\n" +
    '  match_status,\n' +
    '  COUNT(*) AS n\n' +
    'FROM `' + PROJECT_ID + '.' + STOCK_STOCK_TABLE + '`\n' +
    "WHERE qty > 0 AND supplier_pid NOT LIKE r'DELETED\\_%'\n" +
    '  AND supplier_name IN (' + sqlQuoteList(names) + ')\n' +
    'GROUP BY product_type, match_status';
}

function stockProductTypesSQL(names) {
  return 'SELECT\n' +
    '  CASE\n' +
    '    WHEN ' + WHEEL_PATTERN_SQL + '\n' +
    "    THEN 'wheel'\n" +
    '    WHEN ' + NO_DESC_PATTERN_SQL + '\n' +
    "    THEN 'tyre'\n" +
    '    WHEN ' + TYRE_PATTERN_SQL + '\n' +
    "    THEN 'tyre'\n" +
    "    ELSE 'other'\n" +
    '  END AS product_type,\n' +
    '  match_status,\n' +
    '  COUNT(*) AS n\n' +
    'FROM `' + PROJECT_ID + '.' + STOCK_STOCK_TABLE + '`\n' +
    "WHERE qty > 0 AND supplier_pid NOT LIKE r'DELETED\\_%'\n" +
    '  AND supplier_name IN (' + sqlQuoteList(names) + ')\n' +
    'GROUP BY product_type, match_status';
}

async function fetchProductTypeBreakdown(keys) {
  const names = stockRowsForKeys(keys);
  const empty = { wheel: { matched: 0, unmatched: 0, total: 0 }, tyre: { matched: 0, unmatched: 0, total: 0 }, other: { matched: 0, unmatched: 0, total: 0 } };
  if (!names.length) return empty;
  const useExactMatch = keys.length === 1 && keys[0] === 'chapelcorner';
  const rows = await runBQQuery(useExactMatch ? stockProductTypesSQLExactMatch(names) : stockProductTypesSQL(names));
  const out = { wheel: { matched: 0, unmatched: 0, total: 0 }, tyre: { matched: 0, unmatched: 0, total: 0 }, other: { matched: 0, unmatched: 0, total: 0 } };
  for (const r of rows) {
    const type = (r.product_type === 'wheel' || r.product_type === 'other') ? r.product_type : 'tyre';
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

