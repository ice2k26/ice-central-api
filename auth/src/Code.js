/**
 * ICE Auth — Google sign-in broker.
 *
 * Deployed as: execute as USER_ACCESSING, access ANYONE (Google sign-in required).
 * Google forces the visitor to sign in before this doGet runs, so
 * Session.getActiveUser().getEmail() is the verified identity of the visitor.
 * We mint an HMAC-signed bearer token and bounce the browser back to the
 * frontend with it in the URL hash. The API project (separate deployment,
 * executes as owner) verifies the same HMAC using the shared SECRET.
 *
 * SECRET lives in Script Properties (key "SECRET"), NOT in source: because
 * this web app executes as USER_ACCESSING, the script file must be shared
 * read-only with every visitor ("anyone with link — viewer"), so anything
 * in source is world-readable. Script properties are not visible to viewers.
 * The api/ project keeps the same value in its git-ignored Secret.js.
 */

var TOKEN_TTL_DAYS = 30;

// Redirect allowlist — the token is only ever handed to a trusted origin.
// Every project lives on its own {slug}.designthinking.lk subdomain, so we
// allow ANY https subdomain of designthinking.lk (wildcard) — new projects work
// with no allowlist edit — plus localhost for dev. See isAllowedRedirect_.

function doGet(e) {
  var params = (e && e.parameter) || {};
  var redirect = params.redirect || '';

  var email = '';
  try {
    email = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  } catch (err) {
    email = '';
  }

  if (!email) {
    return renderMessage_(
      'Sign-in failed',
      'Google did not provide your email address. Please try again in a regular browser window (not incognito), or contact the organizers.'
    );
  }

  if (!isAllowedRedirect_(redirect)) {
    return renderMessage_(
      'Invalid redirect',
      'This sign-in link points to an unrecognized site (' + escapeHtml_(redirect) + '). Please start again from the official ICE workshop website.'
    );
  }

  var token = mintToken_(email);
  var sep = redirect.indexOf('#') === -1 ? '#' : '&';
  var target = redirect + sep + 'icetoken=' + encodeURIComponent(token);

  // "Use a different account" — Apps Script has no in-page account picker;
  // Session identity follows the browser's active Google session. Bounce
  // through Google's AccountChooser, which shows every account signed in on
  // this browser and continues back to this exec URL (re-running doGet as the
  // chosen account). continue= must stay on a Google-owned host, which the
  // script.google.com exec URL satisfies.
  var self = ScriptApp.getService().getUrl() +
    '?redirect=' + encodeURIComponent(redirect);
  var switchUrl = 'https://accounts.google.com/AccountChooser?continue=' +
    encodeURIComponent(self);

  // The sandboxed iframe blocks scripted top-level navigation, so a
  // user-gesture link (target=_top) is the only reliable way back.
  var page = HtmlService.createHtmlOutput(pageShell_(
    'Signing in…',
    '<h1>You&#39;re signed in</h1>' +
    '<p class="body"><b>' + escapeHtml_(email) + '</b></p>' +
    '<a class="btn" href="' + target.replace(/"/g, '&quot;') + '" target="_top">Continue</a>' +
    '<p class="muted"><a href="' + switchUrl.replace(/"/g, '&quot;') + '" target="_top">Use a different account</a></p>'
  ));
  page.setTitle('ICE — Signing in');
  return page;
}

/**
 * Shared page chrome mirroring the frontend theme (web/css/theme.css).
 * The app's light/dark toggle lives in its own localStorage, unreachable
 * from this Google-hosted origin — prefers-color-scheme is the proxy.
 */
function pageShell_(title, inner) {
  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + escapeHtml_(title) + '</title>' +
    '<style>' +
    ':root{--accent:#6100FF;--bg:#FFFFFF;--text:#0E0F11;--text-body:#5E6875;--text-muted:#838D95}' +
    '@media(prefers-color-scheme:dark){:root{--accent:#00D7EE;--bg:#121316;--text:#F2F4F7;--text-body:#B7BEC8;--text-muted:#8A939D}}' +
    'html,body{height:100%}' +
    'body{margin:0;display:flex;align-items:center;justify-content:center;background:var(--bg);color:var(--text);' +
    'font-family:"neue-haas-grotesk-text","Helvetica Neue",-apple-system,"Segoe UI",sans-serif;' +
    '-webkit-font-smoothing:antialiased;text-align:center}' +
    'main{padding:24px;max-width:26rem}' +
    '.mark{width:56px;height:56px;border-radius:14px;margin-bottom:20px}' +
    'h1{font-family:"neue-haas-grotesk-display","Helvetica Neue",-apple-system,"Segoe UI",sans-serif;' +
    'font-size:1.35rem;font-weight:600;letter-spacing:-0.01em;margin:0 0 6px}' +
    'p{margin:0 0 8px;font-size:0.95rem;line-height:1.5}' +
    '.body{color:var(--text-body)}.muted{color:var(--text-muted);font-size:0.85rem;margin-top:16px}' +
    'a{color:var(--accent);text-decoration:none;font-weight:500}a:hover{text-decoration:underline}' +
    '.btn{display:inline-flex;align-items:center;justify-content:center;margin-top:20px;' +
    'padding:10px 28px;border-radius:999px;font-size:14.5px;font-weight:600;line-height:1.2;color:#fff;' +
    'background:linear-gradient(90deg,#00D7EE 0%,#2E6BF6 55%,#6100FF 100%);' +
    'box-shadow:0 4px 16px -4px rgba(97,0,255,0.45);transition:filter .15s,transform .1s}' +
    '.btn:hover{filter:brightness(1.06);text-decoration:none}.btn:active{transform:translateY(1px)}' +
    '</style></head><body><main>' +
    '<svg class="mark" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" aria-label="ICE">' +
    '<defs><linearGradient id="g" x1="1" y1="0" x2="0" y2="1">' +
    '<stop offset="0" stop-color="#00D7EE"/><stop offset="1" stop-color="#6100FF"/>' +
    '</linearGradient></defs>' +
    '<rect width="64" height="64" rx="16" fill="url(#g)"/>' +
    '<text x="32" y="42" font-family="Helvetica, Arial, sans-serif" font-size="26" font-weight="800" fill="#fff" text-anchor="middle">ICE</text>' +
    '</svg>' +
    inner +
    '</main></body></html>';
}

function isAllowedRedirect_(url) {
  if (!url) return false;
  // Any https single-level subdomain of designthinking.lk. The trailing
  // (?:[/?#]|$) is what makes this safe against look-alikes: it rejects
  // https://designthinking.lk.evil.com/ and https://x.designthinking.lk.evil.com/
  // while allowing https://ice2027.designthinking.lk[/...|?...|#...].
  if (/^https:\/\/[a-z0-9-]+\.designthinking\.lk(?:[/?#]|$)/i.test(url)) return true;
  // Local dev.
  if (url.indexOf('http://localhost:') === 0 || url.indexOf('http://127.0.0.1:') === 0) return true;
  return false;
}

/** token = base64url("email|expiryMillis") + "." + base64url(hmac) */
function mintToken_(email) {
  var secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret) throw new Error('SECRET script property is not set');
  var expiry = Date.now() + TOKEN_TTL_DAYS * 24 * 60 * 60 * 1000;
  var payload = email + '|' + expiry;
  var sig = Utilities.computeHmacSha256Signature(payload, secret);
  return b64url_(Utilities.newBlob(payload).getBytes()) + '.' + b64url_(sig);
}

function b64url_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function renderMessage_(title, body) {
  return HtmlService.createHtmlOutput(pageShell_(
    title,
    '<h1>' + escapeHtml_(title) + '</h1><p class="body">' + body + '</p>'
  ));
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
