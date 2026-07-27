/**
 * iceApplePass — HTTP Cloud Function (Gen 2). Two jobs:
 *
 *   1. Issue a signed .pkpass for an ICE member (GET ?at=<token>).
 *   2. Host the Apple PassKit web service so installed passes update live
 *      (register / unregister / list-updated / latest-pass / log) and get
 *      pushed via APNs when their data changes.
 *
 * Live-data source of truth is the ICE Apps Script API — this function pulls
 * a member's current fields via the `wallet_fields` action (HMAC-signed,
 * shared secret ICE_APPLE_HMAC). It stores only device push registrations +
 * a per-pass content hash in Firestore.
 *
 * Serial number scheme:  <projectId>__<userId>   (matches the Google object id
 * tail) so a serial alone recovers both the project and the user.
 *
 * Secrets (Secret Manager, --set-secrets)
 *   ICE_PASS_P12            — .p12 bytes (base64 when mounted)
 *   ICE_PASS_P12_PASSWORD   — .p12 export password
 *   ICE_APPLE_HMAC          — shared with the ICE API (Script Property
 *                             WALLET_APPLE_HMAC). Signs pass tokens, the
 *                             per-pass authenticationToken, wallet_fields
 *                             calls, and the /internal/refresh key.
 * Env (--set-env-vars)
 *   APPLE_TEAM_ID           — 10-char team id (9B4A86RFZ8)
 *   ICE_API_URL             — the ICE Apps Script /exec URL
 *   WEB_SERVICE_URL         — optional; overrides the req-derived pass
 *                             webServiceURL (must be the function's root URL)
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const http2 = require('http2');
const forge = require('node-forge');
const functions = require('@google-cloud/functions-framework');
const { PKPass } = require('passkit-generator');
const { Firestore } = require('@google-cloud/firestore');

const PASS_TYPE_IDENTIFIER = 'pass.lk.designthinking.member';
const ORG_NAME = 'ICE Design Thinking';
const PASS_DESCRIPTION = 'ICE member card';
const BACKGROUND_COLOR = 'rgb(97, 0, 255)';   // #6100FF
const FOREGROUND_COLOR = 'rgb(255, 255, 255)';
const LABEL_COLOR      = 'rgb(226, 214, 255)';
const APNS_HOST = 'https://api.push.apple.com:443';

const db = new Firestore();
const REGS = 'ice_pass_registrations';   // doc: <deviceLibId>::<serial>
const STATE = 'ice_pass_state';          // doc: <serial> → { hash, lastUpdated }

// ── certificates (lazy, cached across warm invocations) ──
let _certs = null;
function loadCerts() {
  if (_certs) return _certs;
  const p12Base64 = (process.env.ICE_PASS_P12 || '').trim();
  const p12Password = (process.env.ICE_PASS_P12_PASSWORD || '').trim();
  if (!p12Base64 || !p12Password) throw new Error('Missing ICE_PASS_P12 / ICE_PASS_P12_PASSWORD');

  const p12Buffer = Buffer.from(p12Base64, 'base64');
  const p12Asn1 = forge.asn1.fromDer(p12Buffer.toString('binary'));
  const p12 = forge.pkcs12.pkcs12FromAsn1(p12Asn1, p12Password);

  const certBags = p12.getBags({ bagType: forge.pki.oids.certBag });
  const certBag = (certBags[forge.pki.oids.certBag] || [])[0];
  if (!certBag) throw new Error('No certificate found inside .p12');
  const signerCert = forge.pki.certificateToPem(certBag.cert);

  const shrouded = p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag });
  const plain    = p12.getBags({ bagType: forge.pki.oids.keyBag });
  const keyBag = (shrouded[forge.pki.oids.pkcs8ShroudedKeyBag] || [])[0]
              || (plain[forge.pki.oids.keyBag] || [])[0];
  if (!keyBag) throw new Error('No private key found inside .p12');
  const signerKey = forge.pki.privateKeyToPem(keyBag.key);

  const wwdr = fs.readFileSync(path.join(__dirname, 'assets', 'wwdr.pem'), 'utf8');
  _certs = { wwdr, signerCert, signerKey };
  return _certs;
}

// ── crypto helpers over the shared HMAC secret ──
function hmacSecret() {
  const s = (process.env.ICE_APPLE_HMAC || '').trim();
  if (!s) throw new Error('Missing ICE_APPLE_HMAC');
  return s;
}
function b64url(buf) { return Buffer.from(buf).toString('base64url'); }

/** Verify the pass-issue token minted by the ICE API (walletSignAppleToken_):
 *  base64url(json{uid,pid,exp}).base64url(HMAC-SHA256(payloadB64, secret)). */
function verifyPassToken(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2) return null;
  const expected = crypto.createHmac('sha256', hmacSecret()).update(parts[0]).digest('base64url');
  if (!safeEq(parts[1], expected)) return null;
  try {
    const p = JSON.parse(Buffer.from(parts[0], 'base64url').toString('utf8'));
    if (!p.uid || !p.pid) return null;
    if (typeof p.exp === 'number' && p.exp < Date.now()) return null;
    return p;
  } catch (e) { return null; }
}

function safeEq(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(Buffer.from(a), Buffer.from(b));
}

/** Per-pass authenticationToken — deterministic HMAC of the serial, so the
 *  web service can verify `Authorization: ApplePass <token>` with no storage. */
function passAuthToken(serial) {
  return crypto.createHmac('sha256', hmacSecret()).update('auth:' + serial).digest('hex');
}
function checkPassAuth(req, serial) {
  const h = req.get('authorization') || '';
  const m = h.match(/^ApplePass\s+(.+)$/i);
  return !!m && safeEq(m[1], passAuthToken(serial));
}

function parseSerial(serial) {
  const i = String(serial).indexOf('__');
  if (i === -1) return null;
  return { pid: serial.slice(0, i), uid: serial.slice(i + 2) };
}

// ── live fields from the ICE API ──
async function fetchFields(serial) {
  const parsed = parseSerial(serial);
  if (!parsed) throw new Error('bad serial');
  const ts = Date.now();
  const sig = crypto.createHmac('sha256', hmacSecret()).update(serial + '|' + ts).digest('hex');
  const body = JSON.stringify({ action: 'wallet_fields', project: parsed.pid, serial, ts, sig });
  const resp = await fetch(process.env.ICE_API_URL, {
    method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body,
  });
  const data = await resp.json();
  if (!data || data.ok === false || !data.fields) {
    throw new Error('wallet_fields failed: ' + (data && data.message || resp.status));
  }
  return { fields: data.fields, hash: data.hash || '' };
}

// ── build a signed .pkpass ──
function webServiceUrl(req) {
  return (process.env.WEB_SERVICE_URL || ('https://' + req.get('host'))).replace(/\/+$/, '');
}

async function buildPass(req, serial, fields) {
  const teamId = process.env.APPLE_TEAM_ID;
  if (!teamId) throw new Error('APPLE_TEAM_ID missing');
  const certs = loadCerts();
  const parsed = parseSerial(serial);
  const profileUrl = 'https://ice2026.designthinking.lk/#/profile/' + encodeURIComponent(parsed.uid);

  const pass = await PKPass.from({
    model: path.join(__dirname, 'model.pass'),
    certificates: {
      wwdr: certs.wwdr,
      signerCert: certs.signerCert,
      signerKey: certs.signerKey,
      signerKeyPassphrase: (process.env.ICE_PASS_P12_PASSWORD || '').trim(),
    },
  }, {
    passTypeIdentifier: PASS_TYPE_IDENTIFIER,
    teamIdentifier: teamId,
    organizationName: ORG_NAME,
    description: PASS_DESCRIPTION,
    serialNumber: serial,
    backgroundColor: BACKGROUND_COLOR,
    foregroundColor: FOREGROUND_COLOR,
    labelColor: LABEL_COLOR,
    webServiceURL: webServiceUrl(req),
    authenticationToken: passAuthToken(serial),
  });

  pass.type = 'generic';
  // Apple's front has 4 bands top→bottom: header, primary, secondary, auxiliary
  // (nothing renders below auxiliary). To give NOW/UP NEXT and LATEST room for
  // longer text, TEAM|TEAM SCORE ride the top header strip, freeing the two
  // lower bands: NOW|UP NEXT share one row and LATEST gets a full-width row.
  // changeMessage makes iOS show a notification with the new value when the
  // field changes on a live update ("%@" = the new value).
  pass.headerFields.push(
    { key: 'team',  label: 'TEAM',  value: fields.team || 'Unassigned' },
    { key: 'score', label: 'SCORE', value: (fields.score || 0) + ' pts', changeMessage: 'Team score: %@' },
  );
  pass.primaryFields.push({ key: 'name', label: (fields.role || 'Member'), value: fields.name || '' });
  pass.secondaryFields.push(
    { key: 'now',  label: 'NOW',     value: fields.now || '—', changeMessage: 'Now: %@' },
    { key: 'next', label: 'UP NEXT', value: fields.next || '—' },
  );
  pass.auxiliaryFields.push(
    { key: 'note', label: 'LATEST',  value: fields.announcement || '—', changeMessage: '%@' },
  );
  pass.backFields.push(
    { key: 'profile', label: 'Profile', value: profileUrl },
    { key: 'event',   label: 'Event',   value: 'ICE 2026 · 15–17 Aug' },
    { key: 'site',    label: 'Website', value: 'https://ice2026.designthinking.lk' },
  );
  pass.setBarcodes({ format: 'PKBarcodeFormatQR', message: profileUrl, messageEncoding: 'iso-8859-1' });
  return pass;
}

// Build the SHAREABLE project "business card" — a STATIC pass (no webServiceURL,
// no auth token, no APNs). All fields are baked into the signed token by the
// api (apple_project_pass), so nothing is fetched here. Uses its own model dir
// (model.project.pass) whose pass.json omits sharingProhibited → shareable.
async function buildProjectPass(claims) {
  const teamId = process.env.APPLE_TEAM_ID;
  if (!teamId) throw new Error('APPLE_TEAM_ID missing');
  const certs = loadCerts();
  const serial = 'pcard_' + (claims.pid || 'ice') + '__' + (claims.uid || 'x');
  const website = String(claims.pw || '');

  const pass = await PKPass.from({
    model: path.join(__dirname, 'model.project.pass'),
    certificates: {
      wwdr: certs.wwdr,
      signerCert: certs.signerCert,
      signerKey: certs.signerKey,
      signerKeyPassphrase: (process.env.ICE_PASS_P12_PASSWORD || '').trim(),
    },
  }, {
    passTypeIdentifier: PASS_TYPE_IDENTIFIER,
    teamIdentifier: teamId,
    organizationName: ORG_NAME,
    description: 'ICE project card',
    serialNumber: serial,
    backgroundColor: 'rgb(17, 24, 39)',    // #111827 — matches the Google card
    foregroundColor: FOREGROUND_COLOR,
    labelColor: 'rgb(190, 197, 210)',
  });

  pass.type = 'generic';
  pass.primaryFields.push({ key: 'project', value: claims.pn || 'Project' });
  pass.secondaryFields.push(
    { key: 'member', label: 'SHARED BY', value: claims.mn || '' },
    { key: 'team',   label: 'TEAM',      value: claims.tn || '' },
  );
  pass.auxiliaryFields.push({ key: 'about', label: 'ABOUT', value: claims.pd || '—' });
  if (website) {
    // a URL field renders tappable on iOS; the QR opens the same site
    pass.backFields.push({ key: 'site', label: 'Project site', value: website });
  }
  pass.setBarcodes({
    format: 'PKBarcodeFormatQR',
    message: website || 'https://ice2026.designthinking.lk',
    messageEncoding: 'iso-8859-1',
  });
  return pass;
}

// ── APNs (mTLS with the pass signing cert; empty payload) ──
function sendApns(pushTokens) {
  if (!pushTokens.length) return Promise.resolve({ sent: 0, gone: [] });
  const certs = loadCerts();
  const client = http2.connect(APNS_HOST, { cert: certs.signerCert, key: certs.signerKey });
  const gone = [];
  let sent = 0;
  const jobs = pushTokens.map(function (tok) {
    return new Promise(function (resolve) {
      const r = client.request({
        ':method': 'POST', ':path': '/3/device/' + tok,
        'apns-topic': PASS_TYPE_IDENTIFIER, 'apns-push-type': 'background', 'apns-priority': '5',
      });
      let status = 0, data = '';
      r.on('response', function (h) { status = h[':status']; });
      r.on('data', function (d) { data += d; });
      r.on('end', function () {
        if (status === 200) sent++;
        else if (status === 410) gone.push(tok);   // device no longer registered
        else console.warn('APNs', status, data.slice(0, 120));
        resolve();
      });
      r.on('error', function (e) { console.warn('APNs err', e.message); resolve(); });
      r.setTimeout(10000, function () { r.close(); resolve(); });
      r.end(JSON.stringify({}));
    });
  });
  return Promise.all(jobs).then(function () { client.close(); return { sent: sent, gone: gone }; });
}

// ── Firestore helpers ──
function regId(deviceLibId, serial) { return deviceLibId + '::' + serial; }

async function pushTokensForSerial(serial) {
  const snap = await db.collection(REGS).where('serialNumber', '==', serial).get();
  const toks = [];
  snap.forEach(function (d) { const v = d.data(); if (v.pushToken) toks.push(v.pushToken); });
  return toks;
}

async function bumpAndPush(serial, hash) {
  await db.collection(STATE).doc(serial).set({ hash: hash, lastUpdated: Date.now() }, { merge: true });
  const toks = await pushTokensForSerial(serial);
  const res = await sendApns(toks);
  // prune devices APNs reported as gone (410)
  for (const tok of res.gone) {
    const snap = await db.collection(REGS).where('serialNumber', '==', serial).where('pushToken', '==', tok).get();
    for (const doc of snap.docs) await doc.ref.delete();
  }
  return res.sent;
}

// ── routing ──
functions.http('iceApplePass', async (req, res) => {
  try {
    const p = req.path || '/';

    // 1) Issue a pass (browser handoff): GET /?at=<token>
    if (req.method === 'GET' && req.query.at) {
      const claims = verifyPassToken(req.query.at);
      if (!claims) return res.status(403).send('forbidden: invalid token');
      const serial = claims.pid + '__' + claims.uid;
      const { fields, hash } = await fetchFields(serial);
      const pass = await buildPass(req, serial, fields);
      await db.collection(STATE).doc(serial).set({ hash: hash, lastUpdated: Date.now() }, { merge: true });
      res.set('Content-Type', 'application/vnd.apple.pkpass');
      res.set('Content-Disposition', 'attachment; filename="ice_' + serial + '.pkpass"');
      return res.status(200).send(pass.getAsBuffer());
    }

    // 1b) Issue a STATIC project business card: GET /?pat=<token>
    //     Fields are baked into the token; nothing is registered or refreshed.
    if (req.method === 'GET' && req.query.pat) {
      const claims = verifyPassToken(req.query.pat);
      if (!claims || claims.typ !== 'pcard') return res.status(403).send('forbidden: invalid token');
      const pass = await buildProjectPass(claims);
      res.set('Content-Type', 'application/vnd.apple.pkpass');
      res.set('Content-Disposition', 'attachment; filename="ice_project_' + (claims.uid || 'card') + '.pkpass"');
      return res.status(200).send(pass.getAsBuffer());
    }

    // 2) PassKit: register a device for updates
    //    POST /v1/devices/{deviceLibId}/registrations/{passTypeId}/{serial}
    let m = p.match(/^\/v1\/devices\/([^/]+)\/registrations\/([^/]+)\/(.+)$/);
    if (m && (req.method === 'POST' || req.method === 'DELETE')) {
      const deviceLibId = decodeURIComponent(m[1]);
      const serial = decodeURIComponent(m[3]);
      if (!checkPassAuth(req, serial)) return res.status(401).send('unauthorized');
      const ref = db.collection(REGS).doc(regId(deviceLibId, serial));
      if (req.method === 'DELETE') { await ref.delete(); return res.status(200).send('ok'); }
      const pushToken = (req.body && req.body.pushToken) || '';
      if (!pushToken) return res.status(400).send('missing pushToken');
      const existed = (await ref.get()).exists;
      await ref.set({ deviceLibraryIdentifier: deviceLibId, serialNumber: serial,
                      pushToken: pushToken, passTypeIdentifier: PASS_TYPE_IDENTIFIER, updatedAt: Date.now() });
      // seed pass state so passesUpdatedSince has a baseline
      const st = await db.collection(STATE).doc(serial).get();
      if (!st.exists) {
        try { const { hash } = await fetchFields(serial); await db.collection(STATE).doc(serial).set({ hash: hash, lastUpdated: Date.now() }); } catch (e) {}
      }
      return res.status(existed ? 200 : 201).send('ok');
    }

    // 3) PassKit: which of a device's passes changed since a tag
    //    GET /v1/devices/{deviceLibId}/registrations/{passTypeId}[?passesUpdatedSince=tag]
    m = p.match(/^\/v1\/devices\/([^/]+)\/registrations\/([^/]+)\/?$/);
    if (m && req.method === 'GET') {
      const deviceLibId = decodeURIComponent(m[1]);
      const since = Number(req.query.passesUpdatedSince || 0) || 0;
      const snap = await db.collection(REGS).where('deviceLibraryIdentifier', '==', deviceLibId).get();
      const serials = [];
      snap.forEach(function (d) { serials.push(d.data().serialNumber); });
      let maxTag = since;
      const changed = [];
      for (const s of serials) {
        const st = await db.collection(STATE).doc(s).get();
        const lu = st.exists ? (st.data().lastUpdated || 0) : 0;
        if (lu > since) { changed.push(s); if (lu > maxTag) maxTag = lu; }
      }
      if (!changed.length) return res.status(204).send('');
      return res.status(200).json({ lastUpdated: String(maxTag), serialNumbers: changed });
    }

    // 4) PassKit: hand back the latest pass
    //    GET /v1/passes/{passTypeId}/{serial}
    m = p.match(/^\/v1\/passes\/([^/]+)\/(.+)$/);
    if (m && req.method === 'GET') {
      const serial = decodeURIComponent(m[2]);
      if (!checkPassAuth(req, serial)) return res.status(401).send('unauthorized');
      const { fields } = await fetchFields(serial);
      const pass = await buildPass(req, serial, fields);
      res.set('Content-Type', 'application/vnd.apple.pkpass');
      res.set('Last-Modified', new Date().toUTCString());
      return res.status(200).send(pass.getAsBuffer());
    }

    // 5) PassKit: device logs
    if (p === '/v1/log' && req.method === 'POST') {
      console.log('PassKit log:', JSON.stringify((req.body && req.body.logs) || req.body || {}));
      return res.status(200).send('ok');
    }

    // 6) Internal refresh (Cloud Scheduler): recompute all registered serials,
    //    push APNs for any whose content hash changed. Protected by shared key.
    if (p === '/internal/refresh' && req.method === 'POST') {
      if (!safeEq(req.get('x-refresh-key') || '', hmacSecret())) return res.status(403).send('forbidden');
      const snap = await db.collection(REGS).get();
      const serials = {};
      snap.forEach(function (d) { serials[d.data().serialNumber] = true; });
      let pushed = 0, changed = 0;
      for (const serial of Object.keys(serials)) {
        try {
          const { hash } = await fetchFields(serial);
          const st = await db.collection(STATE).doc(serial).get();
          const prev = st.exists ? st.data().hash : '';
          if (hash !== prev) { changed++; pushed += await bumpAndPush(serial, hash); }
        } catch (e) { console.warn('refresh', serial, e.message); }
      }
      return res.status(200).json({ serials: Object.keys(serials).length, changed: changed, pushed: pushed });
    }

    return res.status(404).send('not found');
  } catch (err) {
    console.error('iceApplePass error:', err.stack || err.message);
    return res.status(500).send('error: ' + err.message);
  }
});
