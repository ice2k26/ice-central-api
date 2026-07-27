/**
 * Wallet.js — Google Wallet integration for ICE (generic membership pass).
 * ======================================================================
 *
 * Phase 0 lives here first: create the ICE generic *class* under the
 * (shared, ahlab) Wallet issuer. Later phases add the per-member object
 * signer (`wallet_pass` action) and the live-field PATCH trigger.
 *
 * Reuses ahlab's service account + issuer — so the pattern mirrors
 * ahl-site-appscript/wallet.js, but every symbol here is `WALLET_`-prefixed
 * to avoid colliding with Code.js globals (Apps Script shares one scope).
 *
 * Required Script Properties (copied from the ahlab project):
 *   WALLET_ISSUER_ID       — numeric, e.g. 3388000000023130080
 *   WALLET_SA_EMAIL        — ahl-wallet-issuer@...iam.gserviceaccount.com
 *   WALLET_SA_PRIVATE_KEY  — full PEM (BEGIN/END lines)
 * Set automatically by createIceWalletClass_():
 *   WALLET_CLASS_ID        — <issuerId>.ice_member_v2
 *
 * ── One-off setup ──
 *   1. Set the 3 properties above.
 *   2. Run createIceWalletClass() once from the editor; approve consent.
 *   3. Run iceWalletGetClass() to confirm 200.
 *
 * NOTE: the runnable functions below intentionally have NO trailing
 * underscore — Apps Script hides `_`-suffixed functions from the Run menu.
 */

// v2: created with ONE_USER_ALL_DEVICES (non-shareable). Google forbids
// tightening sharing on a class that already has saved passes, so the locked
// rule ships as a NEW class — v1 (shareable, test installs only) is retired.
var WALLET_CLASS_SUFFIX = 'ice_member_v2';

// The project "business card" — a SHAREABLE, static pass (project + the member
// who shares it). Separate class so it keeps MULTIPLE_HOLDERS (collectable)
// while the member card stays locked. No live refresh — fields are frozen when
// the member mints their card.
var WALLET_PROJECT_CLASS_SUFFIX = 'ice_project_v1';

// ── Config ───────────────────────────────────────────────────
function WALLET_config_() {
  var props = PropertiesService.getScriptProperties();
  var issuerId = props.getProperty('WALLET_ISSUER_ID');
  var saEmail  = props.getProperty('WALLET_SA_EMAIL');
  var saKey    = props.getProperty('WALLET_SA_PRIVATE_KEY');
  if (!issuerId || !saEmail || !saKey) {
    throw new Error('Missing Script Property: WALLET_ISSUER_ID / WALLET_SA_EMAIL / WALLET_SA_PRIVATE_KEY');
  }
  // Apps Script's computeRsaSha256Signature needs real newlines in the PEM;
  // a key pasted from a JSON value carries literal "\n" — normalise it.
  var key = String(saKey).replace(/\\n/g, '\n').replace(/\r\n/g, '\n').trim();
  var classId = props.getProperty('WALLET_CLASS_ID') || (issuerId + '.' + WALLET_CLASS_SUFFIX);
  var projectClassId = props.getProperty('WALLET_PROJECT_CLASS_ID') || (issuerId + '.' + WALLET_PROJECT_CLASS_SUFFIX);
  return { issuerId: issuerId, serviceAccountEmail: saEmail, serviceAccountKey: key,
           classId: classId, projectClassId: projectClassId };
}

// ── base64url helpers (WALLET_-prefixed to avoid collisions) ──
function WALLET_b64url_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}
function WALLET_b64urlStr_(str) {
  return WALLET_b64url_(Utilities.newBlob(str).getBytes());
}

// ── OAuth: mint an access token as the service account ───────
function WALLET_accessToken_() {
  var cfg = WALLET_config_();
  var nowSec = Math.floor(Date.now() / 1000);
  var assertion = {
    iss:   cfg.serviceAccountEmail,
    scope: 'https://www.googleapis.com/auth/wallet_object.issuer',
    aud:   'https://oauth2.googleapis.com/token',
    iat:   nowSec,
    exp:   nowSec + 3600
  };
  var signingInput = WALLET_b64urlStr_(JSON.stringify({ alg: 'RS256', typ: 'JWT' })) +
                     '.' + WALLET_b64urlStr_(JSON.stringify(assertion));
  var sig = Utilities.computeRsaSha256Signature(signingInput, cfg.serviceAccountKey);
  var jwt = signingInput + '.' + WALLET_b64url_(sig);

  var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  });
  var body = JSON.parse(resp.getContentText());
  if (!body.access_token) {
    throw new Error('Token exchange failed (' + resp.getResponseCode() + '): ' + resp.getContentText());
  }
  return body.access_token;
}

// ── Phase 0: create the ICE generic class ────────────────────
// Idempotent-ish: if the class already exists the API returns 409; we
// treat that as success and still persist WALLET_CLASS_ID.
function createIceWalletClass() {
  var cfg = WALLET_config_();
  var classId = cfg.issuerId + '.' + WALLET_CLASS_SUFFIX;
  var token = WALLET_accessToken_();

  // Minimal generic class — all branding (logo, colour, fields) lives on
  // the per-member object, so the class only needs its id. enableSmartTap
  // off (no NFC). notifyPreference lets Wallet surface field-update pushes.
  // ONE_USER_ALL_DEVICES: a member's card syncs across THEIR OWN devices but
  // cannot be shared/added to anyone else's wallet (no "share pass" option) —
  // this is a personal event card, not a transferable ticket.
  var body = {
    id: classId,
    enableSmartTap: false,
    multipleDevicesAndHoldersAllowedStatus: 'ONE_USER_ALL_DEVICES'
  };

  var resp = UrlFetchApp.fetch(
    'https://walletobjects.googleapis.com/walletobjects/v1/genericClass',
    {
      method: 'post',
      headers: { Authorization: 'Bearer ' + token },
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );
  var code = resp.getResponseCode();
  var text = resp.getContentText();
  Logger.log('createIceWalletClass → %s\n%s', code, text);

  if (code === 200 || code === 409) {
    PropertiesService.getScriptProperties().setProperty('WALLET_CLASS_ID', classId);
    Logger.log('WALLET_CLASS_ID set to %s%s', classId,
      code === 409 ? ' (class already existed — OK)' : ' (created)');
    return classId;
  }
  throw new Error('Class insert failed (' + code + '): ' + text);
}

// ── Diagnostics ──────────────────────────────────────────────
function iceWalletGetClass() {
  var cfg = WALLET_config_();
  var token = WALLET_accessToken_();
  var resp = UrlFetchApp.fetch(
    'https://walletobjects.googleapis.com/walletobjects/v1/genericClass/' + cfg.classId,
    { method: 'get', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  Logger.log('Status: %s\nBody: %s', resp.getResponseCode(), resp.getContentText());
}

function iceWalletListClasses() {
  var cfg = WALLET_config_();
  var token = WALLET_accessToken_();
  var resp = UrlFetchApp.fetch(
    'https://walletobjects.googleapis.com/walletobjects/v1/genericClass?issuerId=' + cfg.issuerId,
    { method: 'get', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  Logger.log('Status: %s\nBody: %s', resp.getResponseCode(), resp.getContentText());
}


// =====================================================================
// Phase 1 — live member pass: token, field computation, object builder,
// save-JWT signer, and the list-based refresh trigger.
// =====================================================================

// Object id scheme: <issuerId>.mem_<projectId>__<userId>. The double
// underscore separates projectId (no underscores) from the UUID (hex+dashes,
// no underscores), so the refresh trigger can recover both from Google's own
// object list — no local install-tracking needed.
var WALLET_OBJ_TAG = '.mem_';

function walletObjectId_(cfg, projectId, userId) {
  return cfg.issuerId + WALLET_OBJ_TAG + projectId + '__' + String(userId);
}
function walletParseObjectId_(cfg, fullId) {
  var prefix = cfg.issuerId + WALLET_OBJ_TAG;
  if (String(fullId).indexOf(prefix) !== 0) return null;
  var rest = String(fullId).substring(prefix.length);
  var sep = rest.indexOf('__');
  if (sep === -1) return null;
  return { pid: rest.substring(0, sep), uid: rest.substring(sep + 2) };
}

// ── Wallet QR token (HMAC over the shared SECRET from Secret.js) ──
// Format mirrors verifyToken_ in Code.js: base64url(json).base64url(hmac).
function walletSignToken_(payload) {
  var enc = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(payload)).getBytes()).replace(/=+$/, '');
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(enc, SECRET)).replace(/=+$/, '');
  return enc + '.' + sig;
}
function walletVerifyToken_(token) {
  if (!token || token.indexOf('.') === -1) return null;
  var parts = token.split('.');
  var expected = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(parts[0], SECRET)).replace(/=+$/, '');
  if (expected !== parts[1]) return null;
  try {
    var claims = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
    if (!claims.exp || Date.now() > claims.exp) return null;
    return claims;
  } catch (e) { return null; }
}

// ── Apple: dedicated HMAC secret (NOT the master SECRET — the Apple Cloud
// Function only holds this one, so a function compromise can't forge sessions).
// Script Property WALLET_APPLE_HMAC, mirrored in Secret Manager (ice-apple-hmac).
function walletAppleSecret_() {
  var s = PropertiesService.getScriptProperties().getProperty('WALLET_APPLE_HMAC');
  if (!s) throw new Error('Missing Script Property WALLET_APPLE_HMAC');
  return String(s).trim();
}

// Pass-issue token the #/wallet page hands to the Apple function via ?at=.
function walletSignAppleToken_(payload) {
  var secret = walletAppleSecret_();
  var enc = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(payload)).getBytes()).replace(/=+$/, '');
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(enc, secret)).replace(/=+$/, '');
  return enc + '.' + sig;
}

// Verify a wallet_fields server-to-server call from the Apple function.
function walletVerifyAppleSig_(serial, ts, sig) {
  var n = Number(ts);
  if (!n || Math.abs(Date.now() - n) > 5 * 60 * 1000) return false;   // 5-min window
  var expected = '';
  var bytes = Utilities.computeHmacSha256Signature(serial + '|' + ts, walletAppleSecret_());
  for (var i = 0; i < bytes.length; i++) { var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i]; expected += ('0' + b.toString(16)).slice(-2); }
  return expected === String(sig);
}

function walletFieldsHash_(fields) {
  var s = JSON.stringify([fields.name, fields.role, fields.team, fields.score, fields.now, fields.next, fields.announcement]);
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s)).replace(/=+$/, '');
}

// ── Live field computation (assumes PROJ is set) ──
// Wallet card title. Community roles win: Mentor, then Participant. Only when
// neither is held (an organiser with no community chip) does it read 'Admin'.
function walletRoleLabel_(user) {
  var roles = rolesOf_(user);
  if (roles.indexOf('mentor') !== -1) return 'Mentor';
  if (roles.indexOf('participant') !== -1) return 'Participant';
  if (roles.indexOf('admin') !== -1) return 'Admin';
  return 'Member';
}

function walletNowNext_() {
  try {
    var calId = getConfig_('PROGRAM_CALENDAR_ID_' + PROJ.id, '') || getConfig_('PROGRAM_CALENDAR_ID', '');
    if (!calId) return { now: '', next: '' };
    var cal = CalendarApp.getCalendarById(calId);
    if (!cal) return { now: '', next: '' };
    var now = new Date();
    var evs = cal.getEvents(new Date(now.getTime() - 12 * 3600 * 1000), new Date(now.getTime() + 36 * 3600 * 1000))
      .filter(function (e) { return !e.isAllDayEvent(); })
      .sort(function (a, b) { return a.getStartTime() - b.getStartTime(); });
    var cur = '', nxt = '';
    for (var i = 0; i < evs.length; i++) {
      var s = evs[i].getStartTime(), e = evs[i].getEndTime();
      if (s <= now && now < e) { if (!cur) cur = evs[i].getTitle(); }
      else if (s > now && !nxt) { nxt = evs[i].getTitle(); }
    }
    return { now: cur, next: nxt };
  } catch (err) {
    return { now: '', next: '' };
  }
}

function walletComputeFields_(user) {
  var teams = readTable_('teams');
  var myTeam = null;
  for (var i = 0; i < teams.length; i++) {
    if (parseArr_(teams[i].members).indexOf(user.id) !== -1) { myTeam = teams[i]; break; }
  }
  var nn = walletNowNext_();
  return {
    name: user.name,
    role: walletRoleLabel_(user),
    team: myTeam ? myTeam.name : 'Unassigned',
    score: myTeam ? (Number(myTeam.score) || 0) : 0,
    now: nn.now,
    next: nn.next,
    announcement: walletLatestPush_()
  };
}

// The most recent admin wallet broadcast — shown as the card's LATEST field.
// Guarded: the wallet_pushes tab may not exist until the first broadcast.
function walletLatestPush_() {
  try {
    var rows = readTable_('wallet_pushes');
    if (!rows.length) return '';
    rows.sort(function (a, b) { return String(b.sentAt).localeCompare(String(a.sentAt)); });
    return rows[0].message || '';
  } catch (e) { return ''; }
}

// ── Generic object builder ──
// Order/ids here are referenced by the class card template (patchIceWallet
// ClassTemplate). 'note' is always present so its template field path resolves.
function walletTextModules_(fields) {
  return [
    { id: 'team',  header: 'TEAM',       body: fields.team || 'Unassigned' },
    { id: 'score', header: 'TEAM SCORE', body: (fields.score || 0) + ' pts' },
    { id: 'now',   header: 'NOW',        body: fields.now || '—' },
    { id: 'next',  header: 'UP NEXT',    body: fields.next || '—' },
    { id: 'note',  header: 'LATEST',     body: fields.announcement || '—' }
  ];
}

// Absolute https base for pass URIs. PROJ.siteUrl is stored as a bare host
// (e.g. "ice2026.designthinking.lk"); Google Wallet rejects scheme-less logo
// and link URIs, so force https:// here.
function walletBaseUrl_() {
  var raw = String(PROJ.siteUrl || 'ice2026.designthinking.lk').replace(/\/+$/, '');
  return /^https?:\/\//i.test(raw) ? raw : 'https://' + raw;
}

function walletBuildObject_(user, cfg, fields) {
  var base = walletBaseUrl_();
  var profileUrl = base + '/#/profile/' + encodeURIComponent(user.id);
  return {
    id: walletObjectId_(cfg, PROJ.id, user.id),
    classId: cfg.classId,
    state: 'ACTIVE',
    cardTitle: { defaultValue: { language: 'en-US', value: PROJ.name || 'ICE 2026' } },
    header:    { defaultValue: { language: 'en-US', value: fields.name || user.name } },
    subheader: { defaultValue: { language: 'en-US', value: fields.role || 'Member' } },
    logo: { sourceUri: { uri: base + '/assets/icon-512.png' } },
    hexBackgroundColor: '#6100FF',
    barcode: { type: 'QR_CODE', value: profileUrl },
    textModulesData: walletTextModules_(fields),
    linksModuleData: { uris: [{ uri: profileUrl, description: 'View profile', id: 'profile' }] }
  };
}

// ── Save-JWT (RS256) ──
function walletSignSaveJwt_(passObject, cfg) {
  var nowSec = Math.floor(Date.now() / 1000);
  var payload = {
    iss: cfg.serviceAccountEmail, aud: 'google', typ: 'savetowallet', iat: nowSec,
    payload: { genericObjects: [passObject] }
  };
  var signingInput = WALLET_b64urlStr_(JSON.stringify({ alg: 'RS256', typ: 'JWT' })) +
                     '.' + WALLET_b64urlStr_(JSON.stringify(payload));
  var sig = Utilities.computeRsaSha256Signature(signingInput, cfg.serviceAccountKey);
  return signingInput + '.' + WALLET_b64url_(sig);
}

// Called by the wallet_pass action. Builds the object from live data and
// returns the Add-to-Google-Wallet URL. The object is created server-side
// only when the user actually taps "Save" in Wallet.
function walletBuildSaveUrl_(user) {
  var cfg = WALLET_config_();
  var fields = walletComputeFields_(user);
  var obj = walletBuildObject_(user, cfg, fields);
  return 'https://pay.google.com/gp/v/save/' + walletSignSaveJwt_(obj, cfg);
}

// ── Project "business card" (shareable, static) ──────────────
// Fields for a member's project card: their team's project + their own name.
// Returns null when the member isn't on a team / the team has no project row,
// so callers can refuse to mint a card.
function projectCardFields_(user) {
  var teams = readTable_('teams').slice().sort(function (a, b) {
    return String(a.name || '').localeCompare(String(b.name || ''));
  });
  var slot = -1, team = null;
  for (var i = 0; i < teams.length; i++) {
    if (parseArr_(teams[i].members).indexOf(user.id) !== -1) { team = teams[i]; slot = i; break; }
  }
  if (slot === -1) return null;
  var proj = null;
  readTeamProjects_().forEach(function (p) { if (p.slot === slot) proj = p; });
  if (!proj || !(proj.title || '').trim()) return null;
  return {
    projectName: proj.title || 'Project',
    description: proj.description || '',
    website: /^https?:\/\//i.test(proj.website || '') ? proj.website : (proj.website ? 'https://' + proj.website : ''),
    memberName: user.name || '',
    teamName: team.name || ''
  };
}

var WALLET_PROJ_OBJ_TAG = '.pcard_';
function walletProjectObjectId_(cfg, projectId, userId) {
  return cfg.issuerId + WALLET_PROJ_OBJ_TAG + projectId + '__' + String(userId);
}
function walletParseProjectObjectId_(cfg, fullId) {
  var prefix = cfg.issuerId + WALLET_PROJ_OBJ_TAG;
  if (String(fullId).indexOf(prefix) !== 0) return null;
  var rest = String(fullId).substring(prefix.length);
  var sep = rest.indexOf('__');
  if (sep === -1) return null;
  return { pid: rest.substring(0, sep), uid: rest.substring(sep + 2) };
}
// Stable hash of the project-card fields — drives "did anything change?" both in
// the Google refresh and the Apple field-fetch (project_fields).
function walletProjectFieldsHash_(fields) {
  var s = JSON.stringify([fields.projectName, fields.description, fields.website, fields.memberName, fields.teamName]);
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s)).replace(/=+$/, '');
}

function walletBuildProjectObject_(user, cfg, fields) {
  var base = walletBaseUrl_();
  var obj = {
    id: walletProjectObjectId_(cfg, PROJ.id, user.id),
    classId: cfg.projectClassId,
    state: 'ACTIVE',
    cardTitle:  { defaultValue: { language: 'en-US', value: fields.projectName } },
    header:     { defaultValue: { language: 'en-US', value: fields.memberName } },
    subheader:  { defaultValue: { language: 'en-US', value: fields.teamName || (PROJ.name || 'ICE 2026') } },
    logo: { sourceUri: { uri: base + '/assets/icon-512.png' } },
    hexBackgroundColor: '#111827',
    textModulesData: [{ id: 'about', header: 'ABOUT', body: fields.description || '—' }]
  };
  // link button + QR only when the project has a real site
  if (fields.website) {
    obj.barcode = { type: 'QR_CODE', value: fields.website };
    obj.linksModuleData = { uris: [{ uri: fields.website, description: 'Open project', id: 'site' }] };
  }
  return obj;
}

// Returns the Add-to-Google-Wallet URL for a member's project business card,
// or null when the member has no project yet.
function walletBuildProjectSaveUrl_(user) {
  var cfg = WALLET_config_();
  var fields = projectCardFields_(user);
  if (!fields) return null;
  var obj = walletBuildProjectObject_(user, cfg, fields);
  return 'https://pay.google.com/gp/v/save/' + walletSignSaveJwt_(obj, cfg);
}

// Instantly re-sync the Google project cards of a team's members — called the
// moment a project is edited, so holders don't wait for the 5-min tick. Apple
// holders update on the next scheduler push (~5 min; the api can't send APNs).
// Best-effort: a member who never saved a card just 404s and is skipped.
function walletRefreshProjectForTeam_(team) {
  if (!team) return { patched: 0 };
  var cfg, token;
  try { cfg = WALLET_config_(); token = WALLET_accessToken_(); }
  catch (e) { Logger.log('wallet not configured, skip project refresh: ' + e); return { patched: 0 }; }
  var members = parseArr_(team.members || []);
  var patched = 0;
  for (var i = 0; i < members.length; i++) {
    try {
      var user = rowById_('users', members[i]);
      if (!user) continue;
      var pf = projectCardFields_(user);
      if (!pf) continue;
      var fresh = walletBuildProjectObject_(user, cfg, pf);
      var patch = {
        cardTitle: fresh.cardTitle, header: fresh.header, subheader: fresh.subheader,
        textModulesData: fresh.textModulesData
      };
      if (fresh.barcode) patch.barcode = fresh.barcode;
      if (fresh.linksModuleData) patch.linksModuleData = fresh.linksModuleData;
      var resp = UrlFetchApp.fetch(
        'https://walletobjects.googleapis.com/walletobjects/v1/genericObject/' + fresh.id,
        { method: 'patch', headers: { Authorization: 'Bearer ' + token },
          contentType: 'application/json', payload: JSON.stringify(patch), muteHttpExceptions: true }
      );
      if (resp.getResponseCode() === 200) patched++;
    } catch (e) { Logger.log('project card patch (team) failed: ' + e); }
  }
  return { patched: patched };
}

// ── Refresh trigger: keep every installed pass live ──
// Lists all objects under the ICE class straight from Google (they are the
// source of truth — no local tracking), recomputes each member's live fields,
// and PATCHes only the objects whose visible content changed. A genuinely new
// LATEST announcement also fires a push message so it behaves like the IKEA card.
function walletRefreshTick() {
  var cfg;
  try { cfg = WALLET_config_(); } catch (e) { Logger.log('wallet not configured: ' + e); return; }
  var token = WALLET_accessToken_();
  var objects = walletListAllObjects_(cfg, token);
  var patched = 0, pushed = 0, skipped = 0;

  for (var i = 0; i < objects.length; i++) {
    var obj = objects[i];
    var parsed = walletParseObjectId_(cfg, obj.id);
    if (!parsed) { skipped++; continue; }
    try {
      PROJ = getProject_(parsed.pid);
      if (!PROJ) { skipped++; continue; }
      var user = rowById_('users', parsed.uid);
      if (!user) { skipped++; continue; }

      var fields = walletComputeFields_(user);
      var freshMods = walletTextModules_(fields);
      var curMods = obj.textModulesData || [];
      var curName = (obj.header && obj.header.defaultValue && obj.header.defaultValue.value) || '';
      if (JSON.stringify(curMods) === JSON.stringify(freshMods) && curName === fields.name) { skipped++; continue; }

      var patch = {
        header:    { defaultValue: { language: 'en-US', value: fields.name || user.name } },
        subheader: { defaultValue: { language: 'en-US', value: fields.role || 'Member' } },
        textModulesData: freshMods
      };
      var resp = UrlFetchApp.fetch(
        'https://walletobjects.googleapis.com/walletobjects/v1/genericObject/' + obj.id,
        { method: 'patch', headers: { Authorization: 'Bearer ' + token },
          contentType: 'application/json', payload: JSON.stringify(patch), muteHttpExceptions: true }
      );
      var code = resp.getResponseCode();
      if (code !== 200) { Logger.log('PATCH %s → %s %s', obj.id, code, resp.getContentText().slice(0, 150)); continue; }
      patched++;

      // Push a notification only when the announcement text actually changes.
      var prevNote = walletModBody_(curMods, 'note');
      if (fields.announcement && fields.announcement !== prevNote) {
        walletAddMessage_(obj.id, token, PROJ.name || 'ICE 2026', fields.announcement);
        pushed++;
      }
    } catch (err) {
      Logger.log('refresh %s failed: %s', obj.id, err && err.stack || err);
    }
  }
  // Project business cards: re-sync name / description / link / member from the
  // live project + team data. Shared holders all hold the same object, so one
  // PATCH updates everyone's copy. No push message — a silent field update.
  var pobjects = walletListAllObjects_(cfg, token, cfg.projectClassId);
  var pPatched = 0;
  for (var pi = 0; pi < pobjects.length; pi++) {
    var pobj = pobjects[pi];
    var pparsed = walletParseProjectObjectId_(cfg, pobj.id);
    if (!pparsed) continue;
    try {
      PROJ = getProject_(pparsed.pid);
      if (!PROJ) continue;
      var puser = rowById_('users', pparsed.uid);
      if (!puser) continue;
      var pf = projectCardFields_(puser);
      if (!pf) continue; // member left the team / project cleared — leave as-is
      var fresh = walletBuildProjectObject_(puser, cfg, pf);
      var curTitle = (pobj.cardTitle && pobj.cardTitle.defaultValue && pobj.cardTitle.defaultValue.value) || '';
      var curHead  = (pobj.header && pobj.header.defaultValue && pobj.header.defaultValue.value) || '';
      var curSub   = (pobj.subheader && pobj.subheader.defaultValue && pobj.subheader.defaultValue.value) || '';
      var curAbout = walletModBody_(pobj.textModulesData, 'about');
      var curLink  = (pobj.linksModuleData && pobj.linksModuleData.uris && pobj.linksModuleData.uris[0] && pobj.linksModuleData.uris[0].uri) || '';
      if (curTitle === fresh.cardTitle.defaultValue.value && curHead === fresh.header.defaultValue.value &&
          curSub === fresh.subheader.defaultValue.value && curAbout === fresh.textModulesData[0].body &&
          curLink === ((fresh.linksModuleData && fresh.linksModuleData.uris[0] && fresh.linksModuleData.uris[0].uri) || '')) {
        continue;
      }
      var ppatch = {
        cardTitle: fresh.cardTitle, header: fresh.header, subheader: fresh.subheader,
        textModulesData: fresh.textModulesData
      };
      if (fresh.barcode) ppatch.barcode = fresh.barcode;
      if (fresh.linksModuleData) ppatch.linksModuleData = fresh.linksModuleData;
      var presp = UrlFetchApp.fetch(
        'https://walletobjects.googleapis.com/walletobjects/v1/genericObject/' + pobj.id,
        { method: 'patch', headers: { Authorization: 'Bearer ' + token },
          contentType: 'application/json', payload: JSON.stringify(ppatch), muteHttpExceptions: true }
      );
      if (presp.getResponseCode() === 200) pPatched++;
      else Logger.log('PATCH project %s → %s %s', pobj.id, presp.getResponseCode(), presp.getContentText().slice(0, 150));
    } catch (perr) {
      Logger.log('project refresh %s failed: %s', pobj.id, perr && perr.stack || perr);
    }
  }

  Logger.log('walletRefreshTick: %s patched, %s pushed, %s unchanged (of %s); %s project cards patched (of %s)',
    patched, pushed, skipped, objects.length, pPatched, pobjects.length);
  return { patched: patched, pushed: pushed, total: objects.length, projectPatched: pPatched };
}

function walletModBody_(mods, id) {
  for (var i = 0; i < (mods || []).length; i++) { if (mods[i].id === id) return mods[i].body || ''; }
  return '';
}

function walletListAllObjects_(cfg, token, classId) {
  classId = classId || cfg.classId;
  var out = [], pageToken = '';
  do {
    var url = 'https://walletobjects.googleapis.com/walletobjects/v1/genericObject?classId=' +
      encodeURIComponent(classId) + '&maxResults=100' + (pageToken ? '&token=' + encodeURIComponent(pageToken) : '');
    var resp = UrlFetchApp.fetch(url, { method: 'get', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) { Logger.log('list objects → %s %s', resp.getResponseCode(), resp.getContentText().slice(0, 200)); break; }
    var body = JSON.parse(resp.getContentText());
    (body.resources || []).forEach(function (o) { out.push(o); });
    pageToken = (body.pagination && body.pagination.nextPageToken) || '';
  } while (pageToken);
  return out;
}

function walletAddMessage_(objectId, token, header, body) {
  var msg = {
    message: {
      header: header,
      body: body,
      id: 'ann_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, body)).replace(/[^a-zA-Z0-9]/g, '').slice(0, 24),
      // TEXT_AND_NOTIFY = message on the pass AND a system-tray/lock-screen push
      // notification (vs plain TEXT, which only shows inside Wallet). Google caps
      // notifications at 3 per pass per 24h, so we only NOTIFY on admin broadcasts,
      // never on the silent 5-min field refreshes.
      messageType: 'TEXT_AND_NOTIFY'
    }
  };
  UrlFetchApp.fetch(
    'https://walletobjects.googleapis.com/walletobjects/v1/genericObject/' + objectId + '/addMessage',
    { method: 'post', headers: { Authorization: 'Bearer ' + token },
      contentType: 'application/json', payload: JSON.stringify(msg), muteHttpExceptions: true }
  );
}

// ── One-off: install the 5-minute refresh trigger (run once from editor) ──
function installWalletTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'walletRefreshTick') {
      Logger.log('walletRefreshTick trigger already installed.');
      return;
    }
  }
  ScriptApp.newTrigger('walletRefreshTick').timeBased().everyMinutes(5).create();
  Logger.log('Installed walletRefreshTick every 5 minutes.');
}

// ── Diagnostics: sign a save URL for the first user and log it ──
// Run from the editor, open the logged pay.google.com URL on a phone to
// verify the round trip before wiring up the frontend.
function testWalletSaveUrl() {
  PROJ = getProject_('ice2026');
  var users = readTable_('users');
  if (!users.length) { Logger.log('No users in ice2026 db yet.'); return; }
  // Prefer the admin's own row so the test card is yours (delete it after).
  var admin = ADMIN_EMAILS[0];
  var user = users.filter(function (u) { return String(u.email).toLowerCase() === admin; })[0] || users[0];
  var fields = walletComputeFields_(user);
  Logger.log('User: %s (%s)', user.name, user.id);
  Logger.log('Live fields: %s', JSON.stringify(fields));
  Logger.log(walletBuildSaveUrl_(user));
}

// ── One-off: define the card FACE layout on the class ──
// A generic pass shows only header/subheader/logo/QR on the face by default.
// This maps the object's text modules onto the front as rows:
//   TEAM | TEAM SCORE          (row 1)
//   NOW  | UP NEXT             (row 2)
//   LATEST                     (row 3)
// Run once; Google re-renders every installed pass within a few minutes.
function patchIceWalletClassTemplate() {
  var cfg = WALLET_config_();
  var token = WALLET_accessToken_();
  function row2(a, b) {
    return { twoItems: {
      startItem: { firstValue: { fields: [{ fieldPath: "object.textModulesData['" + a + "']" }] } },
      endItem:   { firstValue: { fields: [{ fieldPath: "object.textModulesData['" + b + "']" }] } }
    } };
  }
  function row1(a) {
    return { oneItem: { item: { firstValue: { fields: [{ fieldPath: "object.textModulesData['" + a + "']" }] } } } };
  }
  var body = {
    classTemplateInfo: {
      cardTemplateOverride: {
        cardRowTemplateInfos: [ row2('team', 'score'), row2('now', 'next'), row1('note') ]
      }
    }
  };
  var resp = UrlFetchApp.fetch(
    'https://walletobjects.googleapis.com/walletobjects/v1/genericClass/' + cfg.classId,
    { method: 'patch', headers: { Authorization: 'Bearer ' + token },
      contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true }
  );
  Logger.log('patchIceWalletClassTemplate → %s\n%s', resp.getResponseCode(), resp.getContentText());
}

// ── Lock class sharing (only works BEFORE any pass is saved) ──
// ONE_USER_ALL_DEVICES removes Google Wallet's "share pass" option. NOTE:
// Google REJECTS this PATCH (400) once holders have saved the class — it
// can't retroactively enforce the constraint. That's why v2 ships as a fresh
// class (WALLET_CLASS_SUFFIX) created locked from the start via
// createIceWalletClass(). This helper stays for a class with zero saves.
function patchIceWalletClassSharing() {
  var cfg = WALLET_config_();
  var token = WALLET_accessToken_();
  var body = { multipleDevicesAndHoldersAllowedStatus: 'ONE_USER_ALL_DEVICES' };
  var resp = UrlFetchApp.fetch(
    'https://walletobjects.googleapis.com/walletobjects/v1/genericClass/' + cfg.classId,
    { method: 'patch', headers: { Authorization: 'Bearer ' + token },
      contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true }
  );
  Logger.log('patchIceWalletClassSharing → %s\n%s', resp.getResponseCode(), resp.getContentText());
}

// ── One-off: create the SHAREABLE project business-card class ──
// MULTIPLE_HOLDERS so a card can be collected by anyone it's shared with.
// Run once; sets WALLET_PROJECT_CLASS_ID. 409 (already exists) is treated as OK.
function createIceProjectClass() {
  var cfg = WALLET_config_();
  var classId = cfg.issuerId + '.' + WALLET_PROJECT_CLASS_SUFFIX;
  var token = WALLET_accessToken_();
  var body = {
    id: classId,
    enableSmartTap: false,
    multipleDevicesAndHoldersAllowedStatus: 'MULTIPLE_HOLDERS'
  };
  var resp = UrlFetchApp.fetch(
    'https://walletobjects.googleapis.com/walletobjects/v1/genericClass',
    { method: 'post', headers: { Authorization: 'Bearer ' + token },
      contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true }
  );
  var code = resp.getResponseCode();
  Logger.log('createIceProjectClass → %s\n%s', code, resp.getContentText());
  if (code === 200 || code === 409) {
    PropertiesService.getScriptProperties().setProperty('WALLET_PROJECT_CLASS_ID', classId);
    Logger.log('WALLET_PROJECT_CLASS_ID set to %s%s', classId, code === 409 ? ' (existed — OK)' : ' (created)');
    return classId;
  }
  throw new Error('Project class insert failed (' + code + '): ' + resp.getContentText());
}

// ── One-off: paint the project card FACE (the ABOUT description row) ──
// cardTitle (project), header (member), subheader (team), logo, link button and
// QR render by default; the text module needs a template row to show. Run once
// after createIceProjectClass().
function patchIceProjectClassTemplate() {
  var cfg = WALLET_config_();
  var token = WALLET_accessToken_();
  var body = {
    classTemplateInfo: {
      cardTemplateOverride: {
        cardRowTemplateInfos: [
          { oneItem: { item: { firstValue: { fields: [{ fieldPath: "object.textModulesData['about']" }] } } } }
        ]
      }
    }
  };
  var resp = UrlFetchApp.fetch(
    'https://walletobjects.googleapis.com/walletobjects/v1/genericClass/' + cfg.projectClassId,
    { method: 'patch', headers: { Authorization: 'Bearer ' + token },
      contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true }
  );
  Logger.log('patchIceProjectClassTemplate → %s\n%s', resp.getResponseCode(), resp.getContentText());
}

// ── Send a test push message to the admin's own pass ──
// Run after saving the pass to your phone — you should get a Wallet notification.
function testWalletPush() {
  PROJ = getProject_('ice2026');
  var cfg = WALLET_config_();
  var users = readTable_('users');
  var admin = ADMIN_EMAILS[0];
  var user = users.filter(function (u) { return String(u.email).toLowerCase() === admin; })[0] || users[0];
  if (!user) { Logger.log('No user.'); return; }
  var token = WALLET_accessToken_();
  var objId = walletObjectId_(cfg, PROJ.id, user.id);
  walletAddMessage_(objId, token, PROJ.name || 'ICE 2026', 'Live test ✅ your ICE card is connected — updates will appear here.');
  Logger.log('Pushed test message to %s', objId);
}
