/**
 * ICE API — JSON backend over Google Sheets, serving MULTIPLE projects.
 *
 * Deployed as: execute as USER_DEPLOYING (owner), access ANYONE_ANONYMOUS.
 * The frontend (static site on GitHub Pages) talks to this endpoint with
 * POST + Content-Type: text/plain (CORS simple request — no preflight).
 *
 * Auth: bearer tokens minted by the sibling "auth" web app, HMAC-signed with
 * the shared SECRET (Secret.js — git-ignored, present in both projects).
 *
 * Storage: a central REGISTRY spreadsheet (auto-created on first use, ID kept
 * in Script Properties) lists every project (workshop instance — ice2026,
 * ice2027, test runs…) and holds a cross-project people directory. Each
 * project row points at its own database spreadsheet + Drive uploads folder,
 * both auto-created on first use. Every request carries a `project` slug;
 * omitting it falls back to DEFAULT_PROJECT so pre-multi-project clients
 * keep working. Images are served via lh3.googleusercontent.com.
 *
 * Scopes: only https://www.googleapis.com/auth/drive.file — the app can touch
 * ONLY the files it created itself. That's why all storage goes through the
 * Sheets/Drive advanced services (SpreadsheetApp/DriveApp would demand the
 * full drive + spreadsheets scopes), and why "add project" always CREATES
 * sheets — an existing spreadsheet can never be linked in.
 */

var ADMIN_EMAILS = ['sankha@ahlab.org'];

var DEFAULT_PROJECT = 'ice2026';
var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
var REGISTRY_NAME = 'ICE Projects Registry';
var PROJECT_SLUG_RE = /^[a-z0-9][a-z0-9-]{1,15}$/;

// The registry spreadsheet's tabs. `projects`: one row per workshop instance
// (per-project config + storage pointers). `directory`: one row per person,
// keyed by the personal email they sign in with — carries their minted
// @designthinking.lk account and a profile snapshot across projects.
var REGISTRY_TABS = {
  projects: ['id', 'name', 'tagline', 'siteUrl', 'status', 'registrationOpen', 'provisionAccounts', 'dbId', 'uploadsFolderId', 'createdAt', 'updatedAt', 'startDate', 'endDate'],
  directory: ['email', 'workEmail', 'name', 'lastProjectId', 'profile', 'updatedAt'],
};

// The project this invocation operates on — resolved from params.project at
// the top of handle_(). A plain global is safe: Apps Script never shares
// globals between concurrent invocations. IDE-run functions must set it too.
var PROJ = null;

var MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
// Intro / pitch videos (≤60s, Full-HD 1920×1080, checked client-side). Capped
// at 32 MB so the single base64 POST (≈43 MB encoded) stays under the Apps
// Script request ceiling (~50 MB).
var MAX_VIDEO_BYTES = 32 * 1024 * 1024;
var CACHE_TTL_SECONDS = 60;

// Workshop Google Workspace: on registration we mint firstname@designthinking.lk
// (a verified secondary domain in the ahlab.org Workspace) inside the /ICE org
// unit, so participants can DM each other in real Google Chat. See README.
var WORKSPACE_DOMAIN = 'designthinking.lk';
var WORKSPACE_OU = '/ICE';

// workEmail = the minted @designthinking.lk address (blank until provisioned).
// invites = the project's allowlist: register refuses emails without a row
// here, and the row fixes the role. Rows outlive registration — "invited vs
// registered" is derived by matching users.email.
var TABLES = {
  users: ['id', 'email', 'name', 'image', 'bio', 'skills', 'affiliation', 'expertise', 'gender', 'links', 'video', 'role', 'createdAt', 'updatedAt', 'workEmail'],
  invites: ['id', 'email', 'role', 'invitedBy', 'createdAt', 'lastSentAt', 'sendCount'],
  // 'score' is appended LAST so existing team rows (8 cols) stay column-aligned.
  teams: ['id', 'name', 'description', 'coverImage', 'lookingFor', 'creatorId', 'members', 'createdAt', 'updatedAt', 'score'],
  team_links: ['id', 'teamId', 'createdBy', 'title', 'url', 'description', 'createdAt'],
  team_posts: ['id', 'teamId', 'createdBy', 'content', 'createdAt'],
  // The six workshop project cards at #/projects. `slot` (0–5) is the display
  // order and also picks the owning team (teams sorted by name, by index), so
  // that team's members — or an admin — may edit title/description/color.
  // New columns are appended LAST so existing rows stay column-aligned. 'video'
  // is a Drive URL; 'description' is the short one, 'fullDescription' the long
  // one; 'website' + 'websiteOk' (reachability flag, '1'/'' set on save).
  team_projects: ['id', 'slot', 'title', 'description', 'color', 'updatedBy', 'updatedAt', 'video', 'fullDescription', 'website', 'websiteOk'],
  messages: ['id', 'senderId', 'receiverId', 'content', 'read', 'createdAt'],
  announcements: ['id', 'title', 'content', 'type', 'authorId', 'isPinned', 'isPublished', 'createdAt', 'updatedAt'],
  // Admin wallet broadcasts — the message shown as the card's LATEST field and
  // pushed to Google/Apple wallets. This tab doubles as the send history.
  wallet_pushes: ['id', 'message', 'sentBy', 'sentAt', 'googleCount', 'appleCount'],
  options: ['category', 'value'],
};

// Seeded into the "options" tab on first read so admins have rows to edit.
// Admins manage form choices by editing that tab directly (category | value).
var DEFAULT_OPTIONS = {
  skill: [
    'UX', 'Interaction Design', 'Study Design', 'Data Science', 'Data Analytics',
    'Machine Learning', 'Hardware', 'Embedded Systems', 'Mobile Apps', 'Web Development',
    'Fundraising', 'Pitch Deck', 'Strategy', 'Business', 'Content Writing',
    'Figma', '3D Printing', 'Electronics', 'Computer Vision', 'Prototyping',
  ],
  gender: ['Female', 'Male', 'Non-binary', 'Prefer not to say'],
};

// Seeded into the "team_projects" tab on first read. Each card is owned by the
// team at the same sorted index (slot); its members (or an admin) may edit it.
var DEFAULT_TEAM_PROJECTS = [
  { title: 'Smart Mobility', description: 'Rethinking how the city moves — accessible transit for everyone.' },
  { title: 'CareConnect', description: 'Bridging patients and caregivers with human-centred health tools.' },
  { title: 'AgriSense', description: 'Data-driven decisions for smallholder farmers.' },
  { title: 'EduPlay', description: 'Learning through play — creative classrooms beyond the textbook.' },
  { title: 'Circular Living', description: 'Designing waste out of everyday life, one household at a time.' },
  { title: 'FinAccess', description: 'Everyday finance for the unbanked and underserved.' },
];

// workEmail is public: it's a workshop chat handle other participants DM.
var USER_PUBLIC_FIELDS = ['id', 'name', 'image', 'bio', 'skills', 'affiliation', 'expertise', 'links', 'video', 'role', 'createdAt', 'workEmail'];

// Fixed workshop teams for admin assignment ("Team A"…"Team F", rows created
// in the teams tab on first use). Per team: 5 participants + 2 mentors — an
// admin assigned to a team occupies a mentor slot.
var TEAM_LETTERS = ['A', 'B', 'C', 'D', 'E', 'F'];
var TEAM_CAP = { participant: 5, mentor: 2 };

// ---------------------------------------------------------------- entrypoints

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  var params = {};
  try {
    params = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'Invalid JSON body' });
  }
  return handle_(params);
}

function handle_(params) {
  try {
    var action = String(params.action || '');
    var fn = ACTIONS[action];
    if (!fn) return json_({ ok: false, error: 'Unknown action: ' + action });

    var slug = String(params.project || DEFAULT_PROJECT).toLowerCase();
    PROJ = getProject_(slug);
    if (!PROJ && action !== 'ping') {
      return json_({ ok: false, error: 'unknown_project', message: 'Unknown project: ' + slug });
    }

    var ctx = { email: null, authEmail: null, user: null, isAdmin: false };
    var authEmail = verifyToken_(params.token);
    if (authEmail) {
      // A person may sign in with their personal email OR their minted
      // @designthinking.lk workspace account. Both resolve to the same primary
      // identity (personal stays primary); authEmail keeps the raw address.
      ctx.authEmail = authEmail;
      ctx.email = canonicalEmail_(authEmail);
      if (PROJ) {
        ctx.user = findUserByEmail_(ctx.email);
        // Directory link missing (e.g. write failed) but the project row still
        // carries the workEmail — match on it and adopt its personal email.
        if (!ctx.user) ctx.user = findUserByWorkEmail_(authEmail);
        if (ctx.user) ctx.email = String(ctx.user.email).toLowerCase();
      }
      ctx.isAdmin = isAdminEmail_(ctx.email) || isAdminEmail_(authEmail) || (ctx.user && hasRole_(ctx.user, 'admin'));
      if (ctx.user) touchPresence_(ctx.user.id); // best-effort online marker
    }

    if (AUTH_REQUIRED[action] && !ctx.email) {
      return json_({ ok: false, error: 'auth', message: 'Please sign in.' });
    }
    if (ADMIN_REQUIRED[action] && !ctx.isAdmin) {
      return json_({ ok: false, error: 'forbidden', message: 'Admins only.' });
    }
    // Every role removed → visitor-level access only ('me' stays open so the
    // frontend can explain). The row keeps all its data; an admin re-adding a
    // role restores everything. Global admins (ADMIN_EMAILS) can't lock
    // themselves out this way.
    if (ctx.user && !ctx.isAdmin && AUTH_REQUIRED[action] && action !== 'me' &&
        rolesOf_(ctx.user).length === 0) {
      return json_({ ok: false, error: 'norole', message: 'Your account has no assigned role. Contact an organizer to restore access.' });
    }

    var result = fn(params, ctx);
    result.ok = result.ok !== false;
    return json_(result);
  } catch (err) {
    console.error('handle_ failed', err && err.stack || err);
    return json_({ ok: false, error: 'server', message: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------- auth

function verifyToken_(token) {
  if (!token || typeof token !== 'string' || token.indexOf('.') === -1) return null;
  try {
    var parts = token.split('.');
    var payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    var expected = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, SECRET)).replace(/=+$/, '');
    if (expected !== parts[1]) return null;
    var pieces = payload.split('|');
    var email = pieces[0];
    var expiry = Number(pieces[1]);
    if (!email || !expiry || Date.now() > expiry) return null;
    return email.toLowerCase();
  } catch (err) {
    return null;
  }
}

function isAdminEmail_(email) {
  return ADMIN_EMAILS.indexOf(String(email).toLowerCase()) !== -1;
}

// -------------------------------------------------------------------- roles
// users.role holds up to MAX_ROLES comma-separated roles: 'admin' plus one of
// 'participant'/'mentor' (those two never coexist). 'none' = every role was
// removed — the row and all data stay, but the person is treated like a
// visitor until an admin assigns a role again. A blank/unknown value counts
// as participant (the historical default). Mirrored by rolesOf() in web/js/app.js.

var PLATFORM_ROLES = ['admin', 'participant', 'mentor'];
var MAX_ROLES = 2;

function rolesOf_(u) {
  if (!u) return [];
  var raw = String(u.role || '').trim().toLowerCase();
  if (raw === 'none') return [];
  if (!raw) return ['participant'];
  var out = [];
  raw.split(',').forEach(function (r) {
    r = r.trim();
    if (PLATFORM_ROLES.indexOf(r) !== -1 && out.indexOf(r) === -1) out.push(r);
  });
  return out.length ? out : ['participant'];
}

function hasRole_(u, role) { return rolesOf_(u).indexOf(role) !== -1; }

/** The sheet-cell value for a role list — 'admin' first, empty list → 'none'. */
function roleValue_(roles) {
  var ordered = PLATFORM_ROLES.filter(function (r) { return roles.indexOf(r) !== -1; });
  return ordered.length ? ordered.join(',') : 'none';
}

// ------------------------------------------------------------------- actions

var AUTH_REQUIRED = {
  me: 1, register: 1, update_profile: 1, upload_image: 1, upload_profile_video: 1, remove_profile_video: 1, check_url: 1, check_email: 1, persona: 1,
  create_team: 1, update_team: 1, delete_team: 1, join_team: 1, leave_team: 1,
  team_link_add: 1, team_link_delete: 1, team_post_add: 1,
  team_project_update: 1, upload_project_video: 1,
  msg_send: 1, msg_inbox: 1, msg_thread: 1,
  ann_create: 1, ann_update: 1, ann_delete: 1,
  admin_add_role: 1, admin_remove_role: 1, admin_delete_user: 1, admin_set_config: 1, admin_provision_email: 1,
  admin_assign_team: 1, admin_set_score: 1, admin_wallet_push: 1, wallet_push_history: 1,
  admin_invite: 1, admin_resend_invite: 1, admin_revoke_invite: 1,
  admin_list_projects: 1, admin_create_project: 1, admin_update_project: 1, admin_user_projects: 1,
  wallet_link: 1, project_wallet_link: 1,
};

var ADMIN_REQUIRED = {
  admin_add_role: 1, admin_remove_role: 1, admin_delete_user: 1, admin_set_config: 1, admin_provision_email: 1,
  admin_assign_team: 1, admin_set_score: 1, admin_wallet_push: 1, wallet_push_history: 1,
  admin_invite: 1, admin_resend_invite: 1, admin_revoke_invite: 1,
  admin_list_projects: 1, admin_create_project: 1, admin_update_project: 1, admin_user_projects: 1,
};

// Mentors and admins may post announcements; edit/delete is author-or-admin.
function canAnnounce_(ctx) {
  return !!(ctx.isAdmin || (ctx.user && hasRole_(ctx.user, 'mentor')));
}

var ACTIONS = {

  ping: function () { return { pong: true, now: new Date().toISOString() }; },

  /** Server-side reachability check for a profile link (no CORS). Returns exists:
   *  true unless the host doesn't resolve or replies 404/410. Bot-blocked hosts
   *  (LinkedIn 999, 401/403) count as existing — the page is there, it just won't
   *  talk to a crawler. Auth-gated so it can't be used as an open proxy. */
  check_url: function (params, ctx) {
    var url = clean_(params.url, 500);
    if (!/^https?:\/\//i.test(url)) return { exists: false, reason: 'format' };
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: 'get',
        followRedirects: true,
        muteHttpExceptions: true,
        validateHttpsCertificates: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; ICE-linkcheck/1.0; +https://ice.designthinking.lk)' },
      });
      var code = resp.getResponseCode();
      return { exists: !(code === 404 || code === 410), status: code };
    } catch (err) {
      // DNS failure, connection refused, timeout, bad certificate → treat as gone.
      return { exists: false, reason: 'unreachable', message: String((err && err.message) || err) };
    }
  },

  /** Is a workshop email free? Used by the register form to show the address the
   *  new account will get. available:true when no Workspace account holds it.
   *  Uses admin.directory.user (Users.get) — no extra scope. */
  check_email: function (params, ctx) {
    var email = clean_(params.email, 120).toLowerCase();
    if (!new RegExp('^[a-z0-9][a-z0-9._-]*@' + WORKSPACE_DOMAIN.replace(/\./g, '\\.') + '$').test(email)) {
      return { available: false, reason: 'format' };
    }
    try {
      if (typeof AdminDirectory === 'undefined') return { available: false, reason: 'unavailable' };
      AdminDirectory.Users.get(email); // throws 404 if the account doesn't exist
      return { available: false, email: email }; // exists → taken
    } catch (err) {
      var m = String((err && err.message) || err);
      if (/not\s*found|404|does not exist|resource/i.test(m)) return { available: true, email: email };
      return { available: false, reason: 'error', message: m };
    }
  },

  /** Live persona blurb for the register/edit card, written by Claude from
   *  whatever profile fields are filled so far. Needs the Script Property
   *  ANTHROPIC_API_KEY; without it returns disabled:true and the frontend
   *  keeps its static copy. Cached by content hash so a form that settles on
   *  the same fields never re-bills. */
  persona: function (params, ctx) {
    var apiKey = getConfig_('ANTHROPIC_API_KEY', '');
    if (!apiKey) return { text: '', disabled: true };
    var fields = {
      name: clean_(params.name, 100),
      role: params.role === 'mentor' ? 'mentor (facilitator)' : 'participant',
      affiliation: clean_(params.affiliation, 200),
      expertise: clean_(params.expertise, 500),
      bio: clean_(params.bio, 2000),
      skills: parseArr_(params.skills).map(function (s) { return clean_(s, 40); }).slice(0, 10),
    };
    if (!fields.name && !fields.affiliation && !fields.expertise && !fields.bio && !fields.skills.length) {
      return { text: '' };
    }
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(fields));
    var cacheKey = 'persona_' + Utilities.base64EncodeWebSafe(digest).slice(0, 40);
    var cache = CacheService.getScriptCache();
    var hit = cache.get(cacheKey);
    if (hit !== null) return { text: hit };
    var text = generatePersona_(apiKey, fields);
    if (text) cache.put(cacheKey, text, 21600); // 6 h
    return { text: text };
  },

  /** One-shot payload for the frontend: directory + teams + announcements. */
  bootstrap: function (params, ctx) {
    var users = readTable_('users').map(function (u) { return projectUser_(u, ctx); });
    var teams = readTable_('teams').map(parseTeam_);
    // Everyone sees published announcements; authors also see their own drafts,
    // and admins see every draft.
    var announcements = readTable_('announcements')
      .filter(function (a) {
        return truthy_(a.isPublished) || ctx.isAdmin || (ctx.user && a.authorId === ctx.user.id);
      })
      .map(parseAnnouncement_);
    var unread = 0;
    if (ctx.user) {
      var myId = ctx.user.id;
      unread = readTable_('messages', true).filter(function (m) {
        return m.receiverId === myId && !truthy_(m.read);
      }).length;
    }
    // Returning person: signed in and known in the cross-project directory but
    // not yet registered in THIS project — hand the frontend their existing
    // work account + last profile so the register form starts prefilled.
    var prefill = null;
    if (ctx.email && !ctx.user) {
      var dir = findDirectory_(ctx.email);
      if (dir) prefill = { workEmail: dir.workEmail || '', profile: safeParse_(dir.profile) };
    }
    // Signed in but not registered: their invite (if any) tells the register
    // card which role this email was pre-assigned. null = not invited.
    var invite = null;
    if (ctx.email && !ctx.user) {
      var invRow = findInviteByEmail_(ctx.email);
      if (invRow) invite = { role: invRow.role === 'mentor' ? 'mentor' : invRow.role === 'admin' ? 'admin' : 'participant' };
    }
    return {
      registrationOpen: PROJ.registrationOpen,
      me: ctx.user ? projectUser_(ctx.user, ctx, true) : null,
      email: ctx.email || undefined,
      invite: invite,
      invites: ctx.isAdmin ? readInvites_() : undefined,
      isAdmin: !!ctx.isAdmin,
      project: projectPublic_(),
      projects: listVisibleProjects_(ctx),
      teamProjects: readTeamProjects_(),
      prefill: prefill,
      // Links to the backing spreadsheet + uploads Drive folder — admins only.
      dbUrl: ctx.isAdmin ? ('https://docs.google.com/spreadsheets/d/' + dbId_() + '/edit') : undefined,
      uploadsUrl: ctx.isAdmin ? ('https://drive.google.com/drive/folders/' + uploadsFolderId_()) : undefined,
      registryUrl: (ctx.email && isAdminEmail_(ctx.email)) ? ('https://docs.google.com/spreadsheets/d/' + registryId_() + '/edit') : undefined,
      unread: unread,
      users: users,
      teams: teams,
      announcements: announcements,
      online: onlineIds_(),
      options: readOptions_(),
    };
  },

  /** Public workshop program: events from a Google Calendar between the
   *  project's startDate/endDate (3-day window fallback). Configure with the
   *  Script Property PROGRAM_CALENDAR_ID (or PROGRAM_CALENDAR_ID_<projectId>
   *  per project) — AND add https://www.googleapis.com/auth/calendar.readonly
   *  to appsscript.json's oauthScopes, then run setup() once in the IDE to
   *  grant it, then redeploy. Until then this returns configured:false and
   *  the frontend keeps its skeleton grid. Cached 5 minutes. */
  program: function (params, ctx) {
    var calId = getConfig_('PROGRAM_CALENDAR_ID_' + PROJ.id, '') || getConfig_('PROGRAM_CALENDAR_ID', '');
    if (!calId) return { configured: false, events: [] };
    var cache = CacheService.getScriptCache();
    var key = 'program_' + PROJ.id;
    var hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (e) { /* refetch */ } }
    var start = PROJ.startDate ? new Date(PROJ.startDate + 'T00:00:00') : new Date();
    if (!PROJ.startDate) start.setHours(0, 0, 0, 0);
    var end = PROJ.endDate ? new Date(PROJ.endDate + 'T23:59:59') : new Date(start.getTime() + 3 * 864e5);
    var out;
    try {
      var cal = CalendarApp.getCalendarById(calId);
      if (!cal) return { configured: false, events: [], message: 'Calendar not accessible: ' + calId };
      // Wall-clock times in the CALENDAR's timezone — the agenda must render
      // identically for every viewer, wherever they open it from.
      var tz = cal.getTimeZone() || Session.getScriptTimeZone();
      var fmt = function (d) { return Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ss"); };
      out = {
        configured: true,
        timeZone: tz,
        events: cal.getEvents(start, end).map(function (ev) {
          return {
            title: ev.getTitle(),
            start: ev.getStartTime().toISOString(),
            end: ev.getEndTime().toISOString(),
            startLocal: fmt(ev.getStartTime()),
            endLocal: fmt(ev.getEndTime()),
            location: ev.getLocation() || '',
            allDay: ev.isAllDayEvent(),
          };
        }),
      };
    } catch (err) {
      // scope not yet granted or bad id — frontend keeps the skeleton
      return { configured: false, events: [], message: String((err && err.message) || err) };
    }
    cache.put(key, JSON.stringify(out), 300);
    return out;
  },

  /** Mint a short-lived (30 min) wallet link for the signed-in user. The
   *  profile card renders this URL as a QR so the user can scan it with their
   *  phone and land on the #/wallet handoff page (which then adds the pass to
   *  Google/Apple Wallet). Auth-gated; the token binds to the user id. */
  wallet_link: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var wt = walletSignToken_({ uid: ctx.user.id, pid: PROJ.id, exp: Date.now() + 30 * 60 * 1000 });
    var base = walletBaseUrl_();
    return { url: base + '/#/wallet?wt=' + encodeURIComponent(wt), token: wt, ttl: 1800 };
  },

  /** Return the Google Wallet "save" URL for a member. Deliberately NOT in
   *  AUTH_REQUIRED — the phone that scanned the QR has no session token, so it
   *  authenticates with the `wt` wallet token instead. Falls back to the
   *  signed-in user when called from within the app. */
  wallet_pass: function (params, ctx) {
    var user = null;
    if (params.wt) {
      var claims = walletVerifyToken_(String(params.wt));
      if (!claims) return { ok: false, error: 'auth', message: 'Wallet link expired — reopen the QR from your profile.' };
      user = rowById_('users', claims.uid);
    } else if (ctx.user) {
      user = ctx.user;
    }
    if (!user) return { ok: false, error: 'auth', message: 'Sign in or scan the QR from your profile card.' };
    try {
      var url = walletBuildSaveUrl_(user);
      return { url: url, google: url };
    } catch (err) {
      return { ok: false, error: 'server', message: 'Wallet pass failed: ' + (err && err.message || err) };
    }
  },

  /** Apple Wallet: return the Cloud Function URL (with a signed ?at= token)
   *  that serves the .pkpass. Same auth model as wallet_pass — `wt` from the
   *  QR handoff, or the signed-in user. */
  apple_pass_link: function (params, ctx) {
    var user = null;
    if (params.wt) {
      var claims = walletVerifyToken_(String(params.wt));
      if (!claims) return { ok: false, error: 'auth', message: 'Wallet link expired — reopen the QR from your profile.' };
      user = rowById_('users', claims.uid);
    } else if (ctx.user) {
      user = ctx.user;
    }
    if (!user) return { ok: false, error: 'auth', message: 'Sign in or scan the QR from your profile card.' };
    var fnUrl = getConfig_('APPLE_PASS_FN_URL', '');
    if (!fnUrl) return { ok: false, error: 'unconfigured', message: 'Apple Wallet is not configured yet.' };
    var token = walletSignAppleToken_({ uid: user.id, pid: PROJ.id, exp: Date.now() + 30 * 60 * 1000 });
    return { url: fnUrl + (fnUrl.indexOf('?') === -1 ? '?' : '&') + 'at=' + encodeURIComponent(token) };
  },

  /** Mint a shareable link to the signed-in member's PROJECT business card.
   *  Longer-lived than the member link (30 days) — it's meant to be handed out
   *  / printed. Refuses when the member has no project yet. */
  project_wallet_link: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    if (!projectCardFields_(ctx.user)) return { ok: false, error: 'noproject', message: 'Join a team with a project first.' };
    var wt = walletSignToken_({ uid: ctx.user.id, pid: PROJ.id, typ: 'pcard', exp: Date.now() + 30 * 24 * 3600 * 1000 });
    var base = walletBaseUrl_();
    return { url: base + '/#/pcard?wt=' + encodeURIComponent(wt), token: wt, ttl: 30 * 24 * 3600 };
  },

  /** Google save URL for a member's project business card. Auth via `wt` (the
   *  shared link) or the signed-in user, mirroring wallet_pass. */
  project_pass: function (params, ctx) {
    var user = null;
    if (params.wt) {
      var claims = walletVerifyToken_(String(params.wt));
      if (!claims || claims.typ !== 'pcard') return { ok: false, error: 'auth', message: 'Card link expired or invalid.' };
      user = rowById_('users', claims.uid);
    } else if (ctx.user) {
      user = ctx.user;
    }
    if (!user) return { ok: false, error: 'auth', message: 'Sign in or scan a shared project card.' };
    try {
      var url = walletBuildProjectSaveUrl_(user);
      if (!url) return { ok: false, error: 'noproject', message: 'This member has no project card yet.' };
      return { url: url, google: url };
    } catch (err) {
      return { ok: false, error: 'server', message: 'Project pass failed: ' + (err && err.message || err) };
    }
  },

  /** Apple static .pkpass link for a member's project business card. The card
   *  fields are baked into the signed token (static — no live service). */
  apple_project_pass: function (params, ctx) {
    var user = null;
    if (params.wt) {
      var claims = walletVerifyToken_(String(params.wt));
      if (!claims || claims.typ !== 'pcard') return { ok: false, error: 'auth', message: 'Card link expired or invalid.' };
      user = rowById_('users', claims.uid);
    } else if (ctx.user) {
      user = ctx.user;
    }
    if (!user) return { ok: false, error: 'auth', message: 'Sign in or scan a shared project card.' };
    var fields = projectCardFields_(user);
    if (!fields) return { ok: false, error: 'noproject', message: 'This member has no project card yet.' };
    var fnUrl = getConfig_('APPLE_PASS_FN_URL', '');
    if (!fnUrl) return { ok: false, error: 'unconfigured', message: 'Apple Wallet is not configured yet.' };
    // fields are fetched live by the function (project_fields) — the token only
    // needs to authorise uid/pid; no card data is baked in.
    var token = walletSignAppleToken_({
      uid: user.id, pid: PROJ.id, typ: 'pcard',
      exp: Date.now() + 30 * 24 * 3600 * 1000
    });
    return { url: fnUrl + (fnUrl.indexOf('?') === -1 ? '?' : '&') + 'pat=' + encodeURIComponent(token) };
  },

  /** Server-to-server (Apple function only): current live fields for a serial.
   *  HMAC-signed with WALLET_APPLE_HMAC; not a browser endpoint. */
  wallet_fields: function (params, ctx) {
    var serial = String(params.serial || '');
    if (!walletVerifyAppleSig_(serial, params.ts, params.sig)) {
      return { ok: false, error: 'forbidden', message: 'bad signature' };
    }
    var sep = serial.indexOf('__');
    if (sep === -1) return { ok: false, error: 'validation', message: 'bad serial' };
    var uid = serial.substring(sep + 2);
    var user = rowById_('users', uid);
    if (!user) return { ok: false, error: 'notfound', message: 'no such member' };
    var fields = walletComputeFields_(user);
    return { fields: fields, hash: walletFieldsHash_(fields) };
  },

  /** Server-to-server (Apple function only): current live fields for a PROJECT
   *  card serial `pcard_<pid>__<uid>`. Lets the static card go live — the Apple
   *  refresh reads this instead of the fields baked at share time. */
  project_fields: function (params, ctx) {
    var serial = String(params.serial || '');
    if (!walletVerifyAppleSig_(serial, params.ts, params.sig)) {
      return { ok: false, error: 'forbidden', message: 'bad signature' };
    }
    if (serial.indexOf('pcard_') !== 0) return { ok: false, error: 'validation', message: 'bad serial' };
    var core = serial.substring('pcard_'.length);
    var sep = core.indexOf('__');
    if (sep === -1) return { ok: false, error: 'validation', message: 'bad serial' };
    var uid = core.substring(sep + 2);
    var user = rowById_('users', uid);
    if (!user) return { ok: false, error: 'notfound', message: 'no such member' };
    var fields = projectCardFields_(user);
    if (!fields) return { ok: false, error: 'noproject', message: 'no project card' };
    return { fields: fields, hash: walletProjectFieldsHash_(fields) };
  },

  /** Short Claude-written description of a skill, for the Skills map's side
   *  panel. Public; cached 6 h per skill so each is billed at most ~4×/day. */
  skill_info: function (params, ctx) {
    var skill = clean_(params.skill, 40);
    if (!skill) return { ok: false, error: 'validation', message: 'Skill required.' };
    var apiKey = getConfig_('ANTHROPIC_API_KEY', '');
    if (!apiKey) return { text: '', disabled: true };
    var key = 'skilldesc_' + skill.toLowerCase().replace(/[^a-z0-9]+/g, '_').slice(0, 60);
    var cache = CacheService.getScriptCache();
    var hit = cache.get(key);
    if (hit !== null) return { text: hit };
    var text = generateSkillBlurb_(apiKey, skill);
    if (text) cache.put(key, text, 21600);
    return { text: text };
  },

  me: function (params, ctx) {
    return {
      registered: !!ctx.user,
      email: ctx.email,
      isAdmin: !!ctx.isAdmin,
      user: ctx.user ? projectUser_(ctx.user, ctx, true) : null,
    };
  },

  register: function (params, ctx) {
    if (ctx.user) return { ok: false, error: 'exists', message: 'You are already registered.' };
    if (!PROJ.registrationOpen && !ctx.isAdmin) {
      return { ok: false, error: 'closed', message: 'Registration is closed.' };
    }
    // Invitation gate: the invites tab is the allowlist, and the invite fixed
    // the role — the card no longer asks. Global admins need no invite.
    var invite = findInviteByEmail_(ctx.email);
    if (!invite && !ctx.isAdmin) {
      return { ok: false, error: 'notinvited', message: 'Registration is by invitation — ask an organizer to invite ' + ctx.email + '.' };
    }
    var first = clean_(params.firstName, 50);
    var last = clean_(params.lastName, 50);
    var name = clean_(params.name, 100) || (first + ' ' + last).trim();
    if (!name) return { ok: false, error: 'validation', message: 'Name is required.' };
    if (!first) { // client sent only a combined name — split it for the email handle
      var parts = name.split(/\s+/);
      first = parts.shift() || '';
      last = parts.join(' ');
    }
    var now = new Date().toISOString();
    // Workshop @designthinking.lk account: returning people (in the directory)
    // keep the one they already have — no duplicate mint, no new password.
    // Otherwise mint one, unless this project has provisioning switched off
    // (test projects). Guarded: registration still succeeds (workEmail just
    // stays blank) if provisioning fails for any reason.
    var dir = findDirectory_(ctx.email);
    var workEmail = '';
    if (dir && dir.workEmail) {
      workEmail = dir.workEmail;
      sendWorkspaceWelcomeBack_(ctx.email, first, workEmail);
    } else if (PROJ.provisionAccounts) {
      workEmail = provisionWorkspaceAccount_(first, last, ctx.email);
    }
    var user = {
      id: Utilities.getUuid(),
      email: ctx.email,
      name: name,
      image: clean_(params.image, 500),
      bio: clean_(params.bio, 2000),
      skills: jsonArr_(params.skills, 30, 40),
      affiliation: clean_(params.affiliation, 200),
      expertise: clean_(params.expertise, 500),
      gender: clean_(params.gender, 30),
      links: jsonArr_(params.links, 10, 300),
      video: clean_(params.video, 300),
      // Pre-assigned by the invite: participant (member/student), mentor
      // (facilitator) or admin (organizer). Global admins always register as admin.
      role: isAdminEmail_(ctx.email) ? 'admin'
          : (invite && invite.role === 'mentor') ? 'mentor'
          : (invite && invite.role === 'admin') ? 'admin'
          : 'participant',
      createdAt: now,
      updatedAt: now,
      workEmail: workEmail,
    };
    appendRow_('users', user);
    upsertDirectory_(ctx.email, {
      workEmail: workEmail,
      name: name,
      lastProjectId: PROJ.id,
      profile: JSON.stringify(profileSnapshot_(user)),
    });
    return { user: projectUser_(user, { isAdmin: true }, true) };
  },

  update_profile: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var patch = { updatedAt: new Date().toISOString() };
    if (params.name !== undefined) {
      var name = clean_(params.name, 100);
      if (!name) return { ok: false, error: 'validation', message: 'Name cannot be empty.' };
      patch.name = name;
    }
    if (params.image !== undefined) patch.image = clean_(params.image, 500);
    if (params.bio !== undefined) patch.bio = clean_(params.bio, 2000);
    if (params.skills !== undefined) patch.skills = jsonArr_(params.skills, 30, 40);
    if (params.affiliation !== undefined) patch.affiliation = clean_(params.affiliation, 200);
    if (params.expertise !== undefined) patch.expertise = clean_(params.expertise, 500);
    if (params.gender !== undefined) patch.gender = clean_(params.gender, 30);
    if (params.links !== undefined) patch.links = jsonArr_(params.links, 10, 300);
    // Intro video is now an uploaded Drive clip (see upload_profile_video); only
    // a Drive URL (or '' to clear) is stored. Legacy YouTube links get dropped.
    if (params.video !== undefined) {
      var pv = clean_(params.video, 300);
      patch.video = (pv && /^https:\/\/(lh3\.googleusercontent\.com|drive\.(google|usercontent\.google)\.com)\//.test(pv)) ? pv : '';
    }
    // Roles are pre-assigned (invite) and admin-managed (admin_add_role /
    // admin_remove_role) — update_profile never touches them.
    updateRowById_('users', ctx.user.id, patch);
    var updated = findUserByEmail_(ctx.email);
    // Keep the cross-project directory snapshot tracking their latest profile.
    upsertDirectory_(ctx.email, {
      name: updated.name,
      lastProjectId: PROJ.id,
      profile: JSON.stringify(profileSnapshot_(updated)),
    });
    return { user: projectUser_(updated, ctx, true) };
  },

  upload_image: function (params, ctx) {
    var data = String(params.data || '');
    var m = data.match(/^data:([-\w.+/]+);base64,(.*)$/);
    var mime = m ? m[1] : String(params.mimeType || 'image/jpeg');
    var b64 = m ? m[2] : data;
    if (!/^image\//.test(mime)) return { ok: false, error: 'validation', message: 'Only images allowed.' };
    var bytes = Utilities.base64Decode(b64);
    if (bytes.length > MAX_UPLOAD_BYTES) return { ok: false, error: 'validation', message: 'Image must be under 5 MB.' };
    var name = (clean_(params.filename, 80) || 'upload') + '-' + Date.now();
    var blob = Utilities.newBlob(bytes, mime, name);
    var file = Drive.Files.create({ name: name, parents: [uploadsFolderId_()] }, blob);
    Drive.Permissions.create({ role: 'reader', type: 'anyone' }, file.id);
    return { url: 'https://lh3.googleusercontent.com/d/' + file.id, fileId: file.id };
  },

  /** Upload a member's intro video to the project's Drive uploads folder and
   *  return its (public) URL. Resolution (1920×1080), length (≤60s) and format
   *  are validated client-side; here we cap the byte size. Runs during a new
   *  registration too (no user row yet), so — like upload_image — it just
   *  stores the file; update_profile persists the URL onto the user row. */
  upload_profile_video: function (params, ctx) {
    var data = String(params.data || '');
    var m = data.match(/^data:([-\w.+/]+);base64,(.*)$/);
    var mime = m ? m[1] : String(params.mimeType || '');
    var b64 = m ? m[2] : data;
    if (!/^video\//.test(mime)) return { ok: false, error: 'validation', message: 'Only video files are allowed.' };
    var bytes = Utilities.base64Decode(b64);
    if (bytes.length > MAX_VIDEO_BYTES) return { ok: false, error: 'validation', message: 'Video must be under 32 MB.' };
    var name = 'intro-' + (clean_(params.filename, 60) || 'video') + '-' + Date.now();
    var blob = Utilities.newBlob(bytes, mime, name);
    var file = Drive.Files.create({ name: name, parents: [uploadsFolderId_()] }, blob);
    Drive.Permissions.create({ role: 'reader', type: 'anyone' }, file.id);
    return { url: 'https://lh3.googleusercontent.com/d/' + file.id, fileId: file.id };
  },

  /** Remove a member's intro clip: delete the Drive file (the one passed by the
   *  client — which may be a just-uploaded, not-yet-saved clip — and the one on
   *  the saved row) and clear the row's video field. Works mid-registration too
   *  (no user row yet): it just deletes the Drive file. */
  remove_profile_video: function (params, ctx) {
    var seen = {};
    [params.url, ctx.user && ctx.user.video].forEach(function (u) {
      var fid = driveFileId_(u);
      if (fid && !seen[fid]) {
        seen[fid] = 1;
        try { Drive.Files.remove(fid); } catch (e) { console.warn('remove video ' + fid + ': ' + ((e && e.message) || e)); }
      }
    });
    if (ctx.user) {
      updateRowById_('users', ctx.user.id, { video: '', updatedAt: new Date().toISOString() });
      var updated = findUserByEmail_(ctx.email);
      upsertDirectory_(ctx.email, {
        name: updated.name,
        lastProjectId: PROJ.id,
        profile: JSON.stringify(profileSnapshot_(updated)),
      });
      return { user: projectUser_(updated, ctx, true) };
    }
    return {};
  },

  // ------------------------------------------------------------------ teams

  create_team: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var name = clean_(params.name, 100);
    if (!name) return { ok: false, error: 'validation', message: 'Team name is required.' };
    var now = new Date().toISOString();
    var team = {
      id: Utilities.getUuid(),
      name: name,
      description: clean_(params.description, 3000),
      coverImage: clean_(params.coverImage, 500),
      lookingFor: clean_(params.lookingFor, 500),
      creatorId: ctx.user.id,
      members: JSON.stringify([ctx.user.id]),
      createdAt: now,
      updatedAt: now,
    };
    appendRow_('teams', team);
    return { team: parseTeam_(team) };
  },

  update_team: function (params, ctx) {
    var team = rowById_('teams', params.teamId);
    if (!team) return { ok: false, error: 'notfound', message: 'Team not found.' };
    if (!canManageTeam_(team, ctx)) return { ok: false, error: 'forbidden', message: 'Only the team creator can edit.' };
    var patch = { updatedAt: new Date().toISOString() };
    if (params.name !== undefined) {
      var name = clean_(params.name, 100);
      if (!name) return { ok: false, error: 'validation', message: 'Team name cannot be empty.' };
      patch.name = name;
    }
    if (params.description !== undefined) patch.description = clean_(params.description, 3000);
    if (params.coverImage !== undefined) patch.coverImage = clean_(params.coverImage, 500);
    if (params.lookingFor !== undefined) patch.lookingFor = clean_(params.lookingFor, 500);
    updateRowById_('teams', team.id, patch);
    return { team: parseTeam_(rowById_('teams', team.id)) };
  },

  /** Edit a #/projects card. Allowed for a member of the owning team (the team
   *  at the same sorted index as the project's slot) or any admin. */
  team_project_update: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var slot = Number(params.slot);
    if (!(slot >= 0 && slot < DEFAULT_TEAM_PROJECTS.length)) return { ok: false, error: 'validation', message: 'Unknown project.' };
    var team = teamForSlot_(slot);
    if (!(ctx.isAdmin || (team && isTeamMember_(team, ctx.user.id)))) {
      return { ok: false, error: 'forbidden', message: 'Only this project\'s team can edit it.' };
    }
    var proj = null;
    readTeamProjects_().forEach(function (p) { if (p.slot === slot) proj = p; });
    if (!proj) return { ok: false, error: 'notfound', message: 'Project not found.' };
    var patch = { updatedBy: ctx.user.id, updatedAt: new Date().toISOString() };
    if (params.title !== undefined) {
      var title = clean_(params.title, 80);
      if (!title) return { ok: false, error: 'validation', message: 'Title cannot be empty.' };
      patch.title = title;
    }
    if (params.description !== undefined) patch.description = clean_(params.description, 300);
    if (params.color !== undefined) {
      var color = clean_(params.color, 10);
      if (!/^pc-[1-6]$/.test(color)) return { ok: false, error: 'validation', message: 'Unknown colour.' };
      patch.color = color;
    }
    // '' clears the pitch video (Remove); a non-empty value must be a Drive URL.
    if (params.video !== undefined) {
      var vid = clean_(params.video, 400);
      if (vid && !/^https:\/\/(lh3\.googleusercontent\.com|drive\.(google|usercontent\.google)\.com)\//.test(vid)) {
        return { ok: false, error: 'validation', message: 'Unrecognised video URL.' };
      }
      patch.video = vid;
    }
    if (params.fullDescription !== undefined) patch.fullDescription = clean_(params.fullDescription, 600);
    // Website: normalise, then curl-check it. A broken URL still saves but is
    // flagged (websiteOk='') so the view can show a warning.
    if (params.website !== undefined) {
      var web = clean_(params.website, 300);
      if (web && !/^https?:\/\//i.test(web)) web = 'https://' + web;
      patch.website = web;
      patch.websiteOk = web ? (urlReachable_(web) ? '1' : '') : '';
    }
    updateRowById_('team_projects', proj.id, patch);
    // push the change to any saved project business cards right away (Google
    // instant; Apple picks it up on its next scheduler tick)
    try { walletRefreshProjectForTeam_(team); } catch (e) { console.warn('project card refresh failed: ' + (e && e.message || e)); }
    return { teamProjects: readTeamProjects_() };
  },

  /** Upload a team's pitch video to the project's Drive uploads folder and save
   *  its (public) URL on the project row. Same team-or-admin guard. Resolution
   *  (1920×1080) and length (≤60s) are validated client-side; here we just cap
   *  the byte size. */
  upload_project_video: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var slot = Number(params.slot);
    if (!(slot >= 0 && slot < DEFAULT_TEAM_PROJECTS.length)) return { ok: false, error: 'validation', message: 'Unknown project.' };
    var team = teamForSlot_(slot);
    if (!(ctx.isAdmin || (team && isTeamMember_(team, ctx.user.id)))) {
      return { ok: false, error: 'forbidden', message: 'Only this project\'s team can edit it.' };
    }
    var data = String(params.data || '');
    var m = data.match(/^data:([-\w.+/]+);base64,(.*)$/);
    var mime = m ? m[1] : String(params.mimeType || '');
    var b64 = m ? m[2] : data;
    if (!/^video\//.test(mime)) return { ok: false, error: 'validation', message: 'Only video files are allowed.' };
    var bytes = Utilities.base64Decode(b64);
    if (bytes.length > MAX_VIDEO_BYTES) return { ok: false, error: 'validation', message: 'Video must be under 32 MB.' };
    var name = 'project-' + slot + '-' + (clean_(params.filename, 60) || 'video') + '-' + Date.now();
    var blob = Utilities.newBlob(bytes, mime, name);
    var file = Drive.Files.create({ name: name, parents: [uploadsFolderId_()] }, blob);
    Drive.Permissions.create({ role: 'reader', type: 'anyone' }, file.id);
    var url = 'https://lh3.googleusercontent.com/d/' + file.id;
    var proj = null;
    readTeamProjects_().forEach(function (p) { if (p.slot === slot) proj = p; });
    if (proj) updateRowById_('team_projects', proj.id, { video: url, updatedBy: ctx.user.id, updatedAt: new Date().toISOString() });
    return { url: url, fileId: file.id, teamProjects: readTeamProjects_() };
  },

  delete_team: function (params, ctx) {
    var team = rowById_('teams', params.teamId);
    if (!team) return { ok: false, error: 'notfound', message: 'Team not found.' };
    if (!canManageTeam_(team, ctx)) return { ok: false, error: 'forbidden', message: 'Only the team creator can delete.' };
    deleteRowsWhere_('team_links', function (r) { return r.teamId === team.id; });
    deleteRowsWhere_('team_posts', function (r) { return r.teamId === team.id; });
    deleteRowById_('teams', team.id);
    return {};
  },

  join_team: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var team = rowById_('teams', params.teamId);
    if (!team) return { ok: false, error: 'notfound', message: 'Team not found.' };
    var members = parseArr_(team.members);
    if (members.indexOf(ctx.user.id) === -1) {
      members.push(ctx.user.id);
      updateRowById_('teams', team.id, { members: JSON.stringify(members), updatedAt: new Date().toISOString() });
    }
    return { team: parseTeam_(rowById_('teams', team.id)) };
  },

  leave_team: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var team = rowById_('teams', params.teamId);
    if (!team) return { ok: false, error: 'notfound', message: 'Team not found.' };
    var members = parseArr_(team.members).filter(function (id) { return id !== ctx.user.id; });
    updateRowById_('teams', team.id, { members: JSON.stringify(members), updatedAt: new Date().toISOString() });
    return { team: parseTeam_(rowById_('teams', team.id)) };
  },

  /** Admin sets a team's live score (shown on the wallet pass). Absolute value
   *  (not a delta). The wallet refresh trigger picks the change up on its next
   *  tick and updates every installed pass for that team's members. */
  admin_set_score: function (params, ctx) {
    var team = rowById_('teams', params.teamId);
    if (!team) return { ok: false, error: 'notfound', message: 'Team not found.' };
    var score = Math.round(Number(params.score));
    if (!isFinite(score)) return { ok: false, error: 'validation', message: 'Score must be a number.' };
    updateRowById_('teams', team.id, { score: score, updatedAt: new Date().toISOString() });
    return { team: parseTeam_(rowById_('teams', team.id)) };
  },

  /** Admin wallet broadcast: log the message to the wallet_pushes history tab,
   *  set it as the card's LATEST field, and push it to every installed pass —
   *  Google (PATCH + addMessage) via the refresh tick, Apple (APNs) via the
   *  Cloud Function's /internal/refresh. Returns per-wallet delivery counts. */
  admin_wallet_push: function (params, ctx) {
    var msg = clean_(params.message, 200);
    if (!msg) return { ok: false, error: 'validation', message: 'Message is required.' };
    gid_('wallet_pushes'); // ensure the tab exists before appending
    var row = {
      id: Utilities.getUuid(), message: msg,
      sentBy: ctx.user ? ctx.user.id : ctx.email,
      sentAt: new Date().toISOString(), googleCount: '', appleCount: '',
    };
    appendRow_('wallet_pushes', row);

    // Google: refresh tick PATCHes the LATEST field + fires the message push.
    // It mutates PROJ as it walks objects across projects, so save/restore it.
    var g = 0, a = 0, savedProj = PROJ;
    try { var gr = walletRefreshTick(); g = (gr && gr.pushed) || 0; } catch (e) { console.error('google push', e); }
    PROJ = savedProj;

    // Apple: trigger the function to recompute + APNs-push changed serials.
    try {
      var fnUrl = getConfig_('APPLE_PASS_FN_URL', '');
      var appleSecret = PropertiesService.getScriptProperties().getProperty('WALLET_APPLE_HMAC');
      if (fnUrl && appleSecret) {
        var resp = UrlFetchApp.fetch(String(fnUrl).replace(/\/+$/, '') + '/internal/refresh', {
          method: 'post', headers: { 'X-Refresh-Key': String(appleSecret).trim() }, muteHttpExceptions: true,
        });
        if (resp.getResponseCode() === 200) { try { a = JSON.parse(resp.getContentText()).pushed || 0; } catch (e2) {} }
      }
    } catch (e) { console.error('apple push', e); }

    try { updateRowById_('wallet_pushes', row.id, { googleCount: String(g), appleCount: String(a) }); } catch (e) {}
    return { ok: true, googleCount: g, appleCount: a };
  },

  /** Admin: recent wallet broadcasts (history), newest first, sender resolved. */
  wallet_push_history: function (params, ctx) {
    var rows;
    try { rows = readTable_('wallet_pushes'); } catch (e) { rows = []; }
    rows.sort(function (a, b) { return String(b.sentAt).localeCompare(String(a.sentAt)); });
    return {
      pushes: rows.slice(0, 50).map(function (r) {
        var u = rowById_('users', r.sentBy);
        return {
          id: r.id, message: r.message, sentAt: r.sentAt,
          sentBy: u ? u.name : (r.sentBy || ''),
          googleCount: r.googleCount, appleCount: r.appleCount,
        };
      }),
    };
  },

  team_detail: function (params, ctx) {
    var team = rowById_('teams', params.teamId);
    if (!team) return { ok: false, error: 'notfound', message: 'Team not found.' };
    var links = readTable_('team_links').filter(function (r) { return r.teamId === team.id; });
    var posts = readTable_('team_posts').filter(function (r) { return r.teamId === team.id; });
    return { team: parseTeam_(team), links: links, posts: posts };
  },

  team_link_add: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var team = rowById_('teams', params.teamId);
    if (!team) return { ok: false, error: 'notfound', message: 'Team not found.' };
    if (!isTeamMember_(team, ctx.user.id) && !ctx.isAdmin) {
      return { ok: false, error: 'forbidden', message: 'Members only.' };
    }
    var url = clean_(params.url, 500);
    if (!/^https?:\/\//.test(url)) return { ok: false, error: 'validation', message: 'A valid link URL is required.' };
    var link = {
      id: Utilities.getUuid(),
      teamId: team.id,
      createdBy: ctx.user.id,
      title: clean_(params.title, 150) || url,
      url: url,
      description: clean_(params.description, 500),
      createdAt: new Date().toISOString(),
    };
    appendRow_('team_links', link);
    return { link: link };
  },

  team_link_delete: function (params, ctx) {
    var link = rowById_('team_links', params.linkId);
    if (!link) return { ok: false, error: 'notfound', message: 'Link not found.' };
    var team = rowById_('teams', link.teamId);
    var mine = ctx.user && link.createdBy === ctx.user.id;
    if (!mine && !canManageTeam_(team, ctx)) return { ok: false, error: 'forbidden', message: 'Not allowed.' };
    deleteRowById_('team_links', link.id);
    return {};
  },

  team_post_add: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var team = rowById_('teams', params.teamId);
    if (!team) return { ok: false, error: 'notfound', message: 'Team not found.' };
    if (!isTeamMember_(team, ctx.user.id) && !ctx.isAdmin) {
      return { ok: false, error: 'forbidden', message: 'Members only.' };
    }
    var content = clean_(params.content, 2000);
    if (!content) return { ok: false, error: 'validation', message: 'Message cannot be empty.' };
    var post = {
      id: Utilities.getUuid(),
      teamId: team.id,
      createdBy: ctx.user.id,
      content: content,
      createdAt: new Date().toISOString(),
    };
    appendRow_('team_posts', post);
    return { post: post };
  },

  // --------------------------------------------------------------- messages

  msg_send: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var to = rowById_('users', params.toId);
    if (!to) return { ok: false, error: 'notfound', message: 'Recipient not found.' };
    var content = clean_(params.content, 2000);
    if (!content) return { ok: false, error: 'validation', message: 'Message cannot be empty.' };
    var msg = {
      id: Utilities.getUuid(),
      senderId: ctx.user.id,
      receiverId: to.id,
      content: content,
      read: 'false',
      createdAt: new Date().toISOString(),
    };
    appendRow_('messages', msg);
    return { message: msg };
  },

  msg_inbox: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var myId = ctx.user.id;
    var mine = readTable_('messages', true).filter(function (m) {
      return m.senderId === myId || m.receiverId === myId;
    });
    var byPeer = {};
    mine.forEach(function (m) {
      var peer = m.senderId === myId ? m.receiverId : m.senderId;
      var e = byPeer[peer] || (byPeer[peer] = { peerId: peer, last: null, unread: 0 });
      if (!e.last || m.createdAt > e.last.createdAt) e.last = m;
      if (m.receiverId === myId && !truthy_(m.read)) e.unread++;
    });
    var conversations = Object.keys(byPeer).map(function (k) { return byPeer[k]; });
    conversations.sort(function (a, b) { return a.last.createdAt < b.last.createdAt ? 1 : -1; });
    return { conversations: conversations };
  },

  msg_thread: function (params, ctx) {
    if (!ctx.user) return { ok: false, error: 'noprofile', message: 'Register first.' };
    var myId = ctx.user.id;
    var peerId = String(params.peerId || '');
    var thread = readTable_('messages', true).filter(function (m) {
      return (m.senderId === myId && m.receiverId === peerId) ||
             (m.senderId === peerId && m.receiverId === myId);
    });
    thread.sort(function (a, b) { return a.createdAt < b.createdAt ? -1 : 1; });
    // mark incoming as read
    var unreadIds = thread.filter(function (m) { return m.receiverId === myId && !truthy_(m.read); })
                          .map(function (m) { return m.id; });
    if (unreadIds.length) markMessagesRead_(unreadIds);
    return { messages: thread };
  },

  // ---------------------------------------------------------- announcements

  ann_create: function (params, ctx) {
    if (!canAnnounce_(ctx)) return { ok: false, error: 'forbidden', message: 'Only mentors and organizers can post announcements.' };
    var title = clean_(params.title, 200);
    var content = clean_(params.content, 5000);
    if (!title || !content) return { ok: false, error: 'validation', message: 'Title and content are required.' };
    var now = new Date().toISOString();
    var ann = {
      id: Utilities.getUuid(),
      title: title,
      content: content,
      type: ['general', 'important', 'urgent'].indexOf(params.type) !== -1 ? params.type : 'general',
      authorId: ctx.user ? ctx.user.id : ctx.email,
      // Only admins may pin (global priority); default published unless saved as a draft.
      isPinned: (ctx.isAdmin && truthy_(params.isPinned)) ? 'true' : 'false',
      isPublished: (params.isPublished === undefined || truthy_(params.isPublished)) ? 'true' : 'false',
      createdAt: now,
      updatedAt: now,
    };
    appendRow_('announcements', ann);
    return { announcement: parseAnnouncement_(ann) };
  },

  ann_update: function (params, ctx) {
    var ann = rowById_('announcements', params.id);
    if (!ann) return { ok: false, error: 'notfound', message: 'Announcement not found.' };
    if (!ctx.isAdmin && !(ctx.user && ann.authorId === ctx.user.id)) {
      return { ok: false, error: 'forbidden', message: 'You can only edit your own announcements.' };
    }
    var patch = { updatedAt: new Date().toISOString() };
    if (params.title !== undefined) patch.title = clean_(params.title, 200);
    if (params.content !== undefined) patch.content = clean_(params.content, 5000);
    if (params.type !== undefined && ['general', 'important', 'urgent'].indexOf(params.type) !== -1) patch.type = params.type;
    if (params.isPinned !== undefined && ctx.isAdmin) patch.isPinned = truthy_(params.isPinned) ? 'true' : 'false';
    if (params.isPublished !== undefined) patch.isPublished = truthy_(params.isPublished) ? 'true' : 'false';
    updateRowById_('announcements', ann.id, patch);
    return { announcement: parseAnnouncement_(rowById_('announcements', ann.id)) };
  },

  ann_delete: function (params, ctx) {
    var ann = rowById_('announcements', params.id);
    if (!ann) return {};
    if (!ctx.isAdmin && !(ctx.user && ann.authorId === ctx.user.id)) {
      return { ok: false, error: 'forbidden', message: 'You can only delete your own announcements.' };
    }
    deleteRowById_('announcements', params.id);
    return {};
  },

  // ------------------------------------------------------------------ admin

  // Role chips: a person holds up to 2 roles — 'admin' plus one of
  // participant/mentor. Removing the last role parks the row as 'none'
  // (visitor-level access, nothing deleted; re-adding a role restores all).
  admin_add_role: function (params, ctx) {
    var user = rowById_('users', params.userId);
    if (!user) return { ok: false, error: 'notfound', message: 'User not found.' };
    var role = String(params.role || '').toLowerCase();
    if (PLATFORM_ROLES.indexOf(role) === -1) {
      return { ok: false, error: 'validation', message: 'Role must be participant, mentor or admin.' };
    }
    var roles = rolesOf_(user);
    if (roles.indexOf(role) !== -1) {
      return { ok: false, error: 'validation', message: user.name + ' already has the ' + role + ' role.' };
    }
    if (roles.length >= MAX_ROLES) {
      return { ok: false, error: 'validation', message: 'At most ' + MAX_ROLES + ' roles per person.' };
    }
    if (role !== 'admin' && (roles.indexOf('participant') !== -1 || roles.indexOf('mentor') !== -1)) {
      return { ok: false, error: 'validation', message: 'Participant and mentor never coexist — remove the current one first.' };
    }
    updateRowById_('users', user.id, { role: roleValue_(roles.concat([role])), updatedAt: new Date().toISOString() });
    return { roles: rolesOf_(rowById_('users', user.id)) };
  },

  admin_remove_role: function (params, ctx) {
    var user = rowById_('users', params.userId);
    if (!user) return { ok: false, error: 'notfound', message: 'User not found.' };
    var role = String(params.role || '').toLowerCase();
    var roles = rolesOf_(user);
    if (roles.indexOf(role) === -1) {
      return { ok: false, error: 'validation', message: user.name + ' does not have the ' + role + ' role.' };
    }
    // Any admin may strip another admin's chip, but never their own — that
    // would break their session mid-flight.
    if (role === 'admin' && ctx.user && ctx.user.id === user.id) {
      return { ok: false, error: 'forbidden', message: 'You cannot remove your own admin role.' };
    }
    updateRowById_('users', user.id, {
      role: roleValue_(roles.filter(function (r) { return r !== role; })),
      updatedAt: new Date().toISOString(),
    });
    // Losing the last community role → drop out of any team (admin-only people
    // don't sit in teams and aren't in the assignable pool). Re-adding a
    // participant/mentor role later lets them be assigned again.
    var updated = rowById_('users', user.id);
    var unassigned = 0;
    if (!isCommunityMember_(updated)) unassigned = removeUserFromTeams_(user.id);
    return { roles: rolesOf_(updated), unassignedFromTeams: unassigned };
  },

  // Per-project purge: removes the person and everything they own IN THIS
  // PROJECT ONLY. The cross-project directory (email <-> workEmail pool) and the
  // @designthinking.lk account are intentionally KEPT so a re-invite reuses the
  // same account; other projects' sheets are separate and untouched.
  admin_delete_user: function (params, ctx) {
    var user = rowById_('users', params.userId);
    if (!user) return { ok: false, error: 'notfound', message: 'User not found.' };
    var uid = user.id;
    var t0 = new Date().getTime();
    console.log('[delete] START project=%s user=%s email=%s name=%s', PROJ.id, uid, user.email, user.name);

    // 1. Drive uploads (profile photo + intro video) in this project's uploads
    //    folder — best-effort, never block the delete. External URLs (YouTube)
    //    produce no fileId and are skipped.
    [['photo', user.image], ['video', user.video]].forEach(function (pair) {
      var fid = driveFileId_(pair[1]);
      if (!fid) { console.log('[delete] %s: no Drive file (skip)', pair[0]); return; }
      try { Drive.Files.remove(fid); console.log('[delete] %s: removed Drive file %s', pair[0], fid); }
      catch (e) { console.error('[delete] %s: Drive remove FAILED %s — %s', pair[0], fid, (e && e.message) || e); }
    });

    // 2. The profile row (also drops their skills from this project's aggregate).
    deleteRowById_('users', uid);
    console.log('[delete] users row removed');

    // 3. Teams: pull from members; delete now-empty teams (+ their posts/links);
    //    reassign a dangling creator to the first remaining member.
    var teamsTouched = 0, teamsDeleted = 0;
    readTable_('teams', true).forEach(function (t) {
      var members = parseArr_(t.members);
      if (members.indexOf(uid) === -1 && t.creatorId !== uid) return;
      var remaining = members.filter(function (m) { return m !== uid; });
      if (remaining.length === 0) {
        deleteRowsWhere_('team_links', function (r) { return r.teamId === t.id; });
        deleteRowsWhere_('team_posts', function (r) { return r.teamId === t.id; });
        deleteRowById_('teams', t.id);
        teamsDeleted++;
        console.log('[delete] team "%s" (%s) now empty — deleted with its posts/links', t.name, t.id);
      } else {
        var patch = { members: JSON.stringify(remaining) };
        if (t.creatorId === uid) { patch.creatorId = remaining[0]; console.log('[delete] team "%s": creator reassigned to %s', t.name, remaining[0]); }
        updateRowById_('teams', t.id, patch);
        teamsTouched++;
      }
    });
    console.log('[delete] teams: %s updated, %s deleted', teamsTouched, teamsDeleted);

    // 4. Authored content in this project. Announcements are KEPT on purpose
    //    (workshop-wide content), so they're not touched here.
    deleteRowsWhere_('messages', function (r) { return r.senderId === uid || r.receiverId === uid; });
    deleteRowsWhere_('team_posts', function (r) { return r.createdBy === uid; });
    deleteRowsWhere_('team_links', function (r) { return r.createdBy === uid; });
    console.log('[delete] authored content (messages/posts/links) removed');

    // 5. Their invite/allowlist entry — remove it so deletion is a full offboard
    //    (otherwise they linger as a pending "Invited" row). Re-inviting brings
    //    them back; the directory pool still makes the workEmail reuse.
    var email = String(user.email || '').toLowerCase();
    deleteRowsWhere_('invites', function (r) { return String(r.email || '').toLowerCase() === email; });
    console.log('[delete] invite/allowlist entry removed (if any)');

    // 6. Directory: KEEP the email <-> workEmail pool (so a re-invite reuses the
    //    same @designthinking.lk account) but WIPE the cached profile snapshot.
    //    bootstrap.prefill.profile reads this — leaving it would repopulate the
    //    register card with the deleted person's old details on re-invite. A
    //    re-registration should start blank, showing only the pre-created work
    //    email. upsertDirectory_ merges, so workEmail/name are preserved.
    try {
      if (findDirectory_(email)) {
        upsertDirectory_(email, { profile: '' });
        console.log('[delete] directory profile snapshot cleared (email/workEmail pool kept)');
      }
    } catch (e) {
      console.error('[delete] directory profile clear FAILED — %s', (e && e.message) || e);
    }

    console.log('[delete] DONE user=%s in %sms — directory pool + @designthinking.lk account kept', uid, (new Date().getTime() - t0));
    return {};
  },

  // Cross-project membership for the admin People view: maps each registered
  // personal email to the list of project ids that person appears in, by
  // scanning every provisioned project's users table. Admin-only; results are
  // cheap on repeat (readTable_ is cached per project).
  admin_user_projects: function (params, ctx) {
    var map = {};
    var savedProj = PROJ;
    try {
      readRegistry_('projects').forEach(function (p) {
        if (!p.dbId) return; // not provisioned yet — no users table
        PROJ = getProject_(p.id, true);
        readTable_('users').forEach(function (u) {
          var key = String(u.email || '').toLowerCase();
          if (!key) return;
          if (!map[key]) map[key] = [];
          if (map[key].indexOf(p.id) === -1) map[key].push(p.id);
        });
      });
    } finally {
      PROJ = savedProj;
    }
    return { memberships: map };
  },

  admin_set_config: function (params, ctx) {
    if (params.registrationOpen !== undefined) {
      var open = truthy_(params.registrationOpen);
      updateRegistryRowByKey_('projects', PROJ.id, {
        registrationOpen: open ? 'true' : 'false',
        updatedAt: new Date().toISOString(),
      });
      PROJ.registrationOpen = open;
    }
    return { registrationOpen: PROJ.registrationOpen };
  },

  // (Re)mint a workshop @designthinking.lk account for a user who has none —
  // for rows that registered before provisioning existed, or where it failed.
  // Reuses the account from a previous project when the directory has one.
  admin_provision_email: function (params, ctx) {
    var u = rowById_('users', params.userId);
    if (!u) return { ok: false, error: 'notfound', message: 'User not found.' };
    if (u.workEmail) return { workEmail: u.workEmail };
    var dir = findDirectory_(u.email);
    var workEmail = (dir && dir.workEmail) || '';
    if (workEmail) {
      var first0 = String(u.name || '').trim().split(/\s+/)[0] || '';
      sendWorkspaceWelcomeBack_(u.email, first0, workEmail);
    } else {
      if (!PROJ.provisionAccounts) {
        return { ok: false, error: 'disabled', message: 'Account provisioning is switched off for this project.' };
      }
      var parts = String(u.name || '').trim().split(/\s+/).filter(Boolean);
      var first = parts.shift() || '';
      var last = parts.join(' ');
      workEmail = provisionWorkspaceAccount_(first, last, u.email);
      if (!workEmail) return { ok: false, error: 'provision', message: 'Could not create the account — check the Admin SDK setup and the execution logs.' };
    }
    updateRowById_('users', u.id, { workEmail: workEmail, updatedAt: new Date().toISOString() });
    upsertDirectory_(u.email, { workEmail: workEmail, name: u.name, lastProjectId: PROJ.id });
    return { workEmail: workEmail };
  },

  // ---------------------------------------------------------------- invites
  // Batch-invite emails as participant or mentor: upsert allowlist rows in the
  // invites tab and email each address an onboarding invitation. Inviting an
  // already-invited email updates its role and re-sends; already-registered
  // emails are skipped and reported.
  admin_invite: function (params, ctx) {
    var role = String(params.role || '').toLowerCase();
    if (role !== 'participant' && role !== 'mentor' && role !== 'admin') {
      return { ok: false, error: 'validation', message: 'Invite role must be participant, mentor or admin.' };
    }
    var raw = Array.isArray(params.emails) ? params.emails : parseArr_(params.emails);
    var emails = [];
    raw.forEach(function (e) {
      e = clean_(e, 120).toLowerCase();
      if (EMAIL_RE.test(e) && emails.indexOf(e) === -1) emails.push(e);
    });
    if (!emails.length) return { ok: false, error: 'validation', message: 'No valid email addresses.' };
    if (emails.length > 50) return { ok: false, error: 'validation', message: 'At most 50 invitations at a time.' };
    var byEmail = {};
    readInvites_(true).forEach(function (i) { byEmail[String(i.email).toLowerCase()] = i; });
    var now = new Date().toISOString();
    var sent = [], alreadyRegistered = [], failed = [];
    emails.forEach(function (email) {
      if (findUserByEmail_(email)) { alreadyRegistered.push(email); return; }
      var mailed = sendInviteEmail_(email, role, ctx.email);
      var existing = byEmail[email];
      if (existing) {
        updateRowById_('invites', existing.id, {
          role: role,
          lastSentAt: mailed ? now : existing.lastSentAt,
          sendCount: String((Number(existing.sendCount) || 0) + (mailed ? 1 : 0)),
        });
      } else {
        appendRow_('invites', {
          id: Utilities.getUuid(), email: email, role: role, invitedBy: ctx.email,
          createdAt: now, lastSentAt: mailed ? now : '', sendCount: mailed ? '1' : '0',
        });
      }
      (mailed ? sent : failed).push(email);
    });
    return { sent: sent, alreadyRegistered: alreadyRegistered, failed: failed, invites: readInvites_(true) };
  },

  admin_resend_invite: function (params, ctx) {
    var inv = rowById_('invites', params.inviteId);
    if (!inv) return { ok: false, error: 'notfound', message: 'Invite not found.' };
    if (!sendInviteEmail_(inv.email, inv.role, ctx.email)) {
      return { ok: false, error: 'mail', message: 'Could not send the email — check the execution logs and the daily mail quota.' };
    }
    updateRowById_('invites', inv.id, {
      lastSentAt: new Date().toISOString(),
      sendCount: String((Number(inv.sendCount) || 0) + 1),
    });
    return { invite: rowById_('invites', inv.id) };
  },

  // Dropping the row closes the allowlist door again; a registered user's row
  // being revoked has no effect on their account.
  admin_revoke_invite: function (params, ctx) {
    var inv = rowById_('invites', params.inviteId);
    if (inv) deleteRowById_('invites', inv.id);
    return {};
  },

  // Put a user into one of the fixed teams (A–F), or pull them out with
  // team: ''. The "Team X" row is created on first assignment. Capacity is
  // enforced here (5 participants + 2 mentors per team); membership is
  // exclusive — assigning removes the user from every other team first.
  admin_assign_team: function (params, ctx) {
    var user = rowById_('users', params.userId);
    if (!user) return { ok: false, error: 'notfound', message: 'User not found.' };
    var letter = clean_(params.team, 1).toUpperCase();
    if (letter && TEAM_LETTERS.indexOf(letter) === -1) {
      return { ok: false, error: 'validation', message: 'Team must be one of ' + TEAM_LETTERS.join(', ') + ' — or empty to unassign.' };
    }
    if (letter && rolesOf_(user).length === 0) {
      return { ok: false, error: 'validation', message: user.name + ' has no assigned role — add one before placing them in a team.' };
    }
    var now = new Date().toISOString();
    var teams = readTable_('teams', true);
    var target = null;
    if (letter) {
      var wanted = ('team ' + letter).toLowerCase();
      teams.forEach(function (t) {
        if (String(t.name || '').trim().toLowerCase() === wanted) target = t;
      });
      if (!target) {
        target = {
          id: Utilities.getUuid(), name: 'Team ' + letter, description: '', coverImage: '',
          lookingFor: '', creatorId: ctx.user ? ctx.user.id : '', members: '[]',
          createdAt: now, updatedAt: now,
        };
        appendRow_('teams', target);
        teams.push(target);
      }
      // Count the target's current slots by role (the assignee excluded, so
      // re-assigning someone already there can never trip the cap).
      var byId = {};
      readTable_('users').forEach(function (u) { byId[u.id] = u; });
      // participant chip → participant slot; mentor (or admin-only) → mentor slot
      var slot = function (u) { return hasRole_(u, 'participant') ? 'participant' : 'mentor'; };
      var used = { participant: 0, mentor: 0 };
      parseArr_(target.members).forEach(function (id) {
        var m = byId[id];
        if (m && m.id !== user.id) used[slot(m)]++;
      });
      var mySlot = slot(user);
      if (used[mySlot] >= TEAM_CAP[mySlot]) {
        return { ok: false, error: 'full', message: target.name + ' already has ' + TEAM_CAP[mySlot] + ' ' + mySlot + 's.' };
      }
    }
    teams.forEach(function (t) {
      var members = parseArr_(t.members);
      var has = members.indexOf(user.id) !== -1;
      if (target && t.id === target.id) {
        if (!has) {
          members.push(user.id);
          updateRowById_('teams', t.id, { members: JSON.stringify(members), updatedAt: now });
        }
      } else if (has) {
        updateRowById_('teams', t.id, {
          members: JSON.stringify(members.filter(function (id) { return id !== user.id; })),
          updatedAt: now,
        });
      }
    });
    return { teams: readTable_('teams', true).map(parseTeam_) };
  },

  // -------------------------------------------------- project management
  // Creating/listing projects is for GLOBAL admins (ADMIN_EMAILS) — a
  // per-project admin must not be able to spawn or enumerate projects.
  // admin_update_project edits the CURRENT project and is open to its admins.

  admin_list_projects: function (params, ctx) {
    if (!isAdminEmail_(ctx.email)) return { ok: false, error: 'forbidden', message: 'Global admins only.' };
    return { projects: readRegistry_('projects', true) };
  },

  // Turnkey project creation: registry row + Google Sheet DB + Drive uploads
  // folder + seeded admin member + Cloudflare subdomain, all done BEFORE we
  // return, so the project is fully usable the moment it can be opened. The row
  // is written first as status:'provisioning' and flipped to 'active' only when
  // every step succeeds — so an accidental client refresh mid-create is safe
  // (the request finishes server-side) and a retry RESUMES rather than
  // duplicating work (each step below is individually idempotent).
  admin_create_project: function (params, ctx) {
    if (!isAdminEmail_(ctx.email)) return { ok: false, error: 'forbidden', message: 'Global admins only.' };
    var id = clean_(params.id, 16).toLowerCase();
    if (!PROJECT_SLUG_RE.test(id)) {
      return { ok: false, error: 'validation', message: 'Project id must be 2–16 chars: lowercase letters, digits, hyphens.' };
    }
    // The subdomain IS the project name — name defaults to the slug.
    var name = clean_(params.name, 60) || id;

    var existing = getProject_(id, true);
    // A finished project is a genuine conflict; a half-built one resumes.
    if (existing && existing.dbId && existing.status !== 'provisioning') {
      return { ok: false, error: 'exists', message: 'A project with that id already exists.' };
    }
    var now = new Date().toISOString();
    if (!existing) {
      appendRegistryRow_('projects', {
        id: id,
        name: name,
        tagline: clean_(params.tagline, 200),
        siteUrl: 'https://' + id + '.designthinking.lk',
        status: 'provisioning',
        registrationOpen: 'true',
        provisionAccounts: 'true', // minting @designthinking.lk accounts is on by default
        dbId: '',
        uploadsFolderId: '',
        createdAt: now,
        updatedAt: now,
      });
    }

    // Everything below is scoped to the NEW project — swap the global PROJ and
    // always restore it (same save/restore pattern as the wallet-refresh path).
    var savedProj = PROJ;
    try {
      PROJ = getProject_(id, true);
      dbId_();               // create the project's Google Sheet DB (writes dbId back)
      uploadsFolderId_();    // create the Drive uploads folder (writes it back)
      seedAdminMember_(ctx); // seed the creating admin as an admin member
      createDnsRecord_(id);  // proxied CNAME so the central-web Worker serves the subdomain
      updateRegistryRowByKey_('projects', id, { status: 'active', updatedAt: new Date().toISOString() });
    } finally {
      PROJ = savedProj;
    }
    return { project: getProject_(id, true) };
  },

  admin_update_project: function (params, ctx) {
    var patch = { updatedAt: new Date().toISOString() };
    if (params.name !== undefined) {
      var name = clean_(params.name, 60);
      if (!name) return { ok: false, error: 'validation', message: 'Project name cannot be empty.' };
      patch.name = name;
    }
    if (params.tagline !== undefined) patch.tagline = clean_(params.tagline, 200);
    if (params.siteUrl !== undefined) patch.siteUrl = clean_(params.siteUrl, 200);
    if (params.status !== undefined) {
      if (['active', 'test', 'archived'].indexOf(params.status) === -1) {
        return { ok: false, error: 'validation', message: 'Status must be active, test or archived.' };
      }
      patch.status = params.status;
    }
    if (params.registrationOpen !== undefined) patch.registrationOpen = truthy_(params.registrationOpen) ? 'true' : 'false';
    if (params.provisionAccounts !== undefined) patch.provisionAccounts = truthy_(params.provisionAccounts) ? 'true' : 'false';
    var dateFields = ['startDate', 'endDate'];
    for (var di = 0; di < dateFields.length; di++) {
      var dk = dateFields[di];
      if (params[dk] !== undefined) {
        var dv = clean_(params[dk], 10);
        if (dv && !/^\d{4}-\d{2}-\d{2}$/.test(dv)) {
          return { ok: false, error: 'validation', message: 'Dates must be YYYY-MM-DD.' };
        }
        patch[dk] = dv;
      }
    }
    if (patch.startDate && patch.endDate && patch.endDate < patch.startDate) {
      return { ok: false, error: 'validation', message: 'End date is before the start date.' };
    }
    updateRegistryRowByKey_('projects', PROJ.id, patch);
    PROJ = getProject_(PROJ.id, true);
    return { project: projectPublic_() };
  },
};

// -------------------------------------------------------------- projections

function projectUser_(u, ctx, includePrivate) {
  var out = {};
  USER_PUBLIC_FIELDS.forEach(function (f) { out[f] = u[f]; });
  out.skills = parseArr_(u.skills);
  out.links = parseArr_(u.links);
  if (includePrivate || (ctx && ctx.isAdmin)) {
    out.email = u.email;
    out.gender = u.gender;
  }
  return out;
}

function parseTeam_(t) {
  var out = {};
  Object.keys(t).forEach(function (k) { out[k] = t[k]; });
  out.members = parseArr_(t.members);
  out.score = Number(t.score) || 0;
  return out;
}

function parseAnnouncement_(a) {
  var out = {};
  Object.keys(a).forEach(function (k) { out[k] = a[k]; });
  out.isPinned = truthy_(a.isPinned);
  out.isPublished = truthy_(a.isPublished);
  return out;
}

/** The current project as the frontend sees it (no storage IDs). */
function projectPublic_() {
  return {
    id: PROJ.id,
    name: PROJ.name,
    tagline: PROJ.tagline,
    siteUrl: PROJ.siteUrl,
    status: PROJ.status,
    registrationOpen: PROJ.registrationOpen,
    provisionAccounts: PROJ.provisionAccounts,
    startDate: PROJ.startDate,
    endDate: PROJ.endDate,
  };
}

/** Projects for the switcher dropdown. Everyone sees active ones; test
 *  projects only show for admins; archived only for global admins. */
function listVisibleProjects_(ctx) {
  var globalAdmin = !!(ctx.email && isAdminEmail_(ctx.email));
  return readRegistry_('projects')
    .filter(function (p) {
      var status = p.status || 'active';
      if (status === 'provisioning') return false; // still being set up — never in the switcher
      if (status === 'active') return true;
      if (status === 'test') return !!ctx.isAdmin || globalAdmin;
      return globalAdmin; // archived
    })
    .map(function (p) { return { id: p.id, name: p.name, tagline: p.tagline, status: p.status || 'active' }; });
}

/** The profile subset stored in the directory for cross-project prefill. */
function profileSnapshot_(u) {
  return {
    name: u.name,
    image: u.image,
    bio: u.bio,
    skills: parseArr_(u.skills),
    affiliation: u.affiliation,
    expertise: u.expertise,
    gender: u.gender,
    links: parseArr_(u.links),
    video: u.video,
  };
}

function canManageTeam_(team, ctx) {
  if (!team) return false;
  if (ctx.isAdmin) return true;
  return !!(ctx.user && team.creatorId === ctx.user.id);
}

function isTeamMember_(team, userId) {
  return parseArr_(team.members).indexOf(userId) !== -1;
}

// ------------------------------------------------------------------ helpers

function clean_(v, maxLen) {
  if (v === undefined || v === null) return '';
  return String(v).trim().slice(0, maxLen);
}

function jsonArr_(v, maxItems, maxLen) {
  var arr = Array.isArray(v) ? v : parseArr_(v);
  arr = arr.map(function (s) { return clean_(s, maxLen); })
           .filter(function (s) { return s.length > 0; })
           .slice(0, maxItems);
  return JSON.stringify(arr);
}

function parseArr_(v) {
  if (Array.isArray(v)) return v;
  if (!v) return [];
  try {
    var parsed = JSON.parse(v);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) { return []; }
}

function truthy_(v) {
  return v === true || v === 'true' || v === 'TRUE' || v === 1 || v === '1';
}

function getConfig_(key, fallback) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return v === null ? fallback : v;
}

function safeParse_(s) {
  if (!s) return null;
  try { return JSON.parse(s); } catch (e) { return null; }
}

// ---------------------------------------------------------------- presence
// Best-effort "online" tracking in CacheService — a per-project map of
// userId → last-seen millis, touched on every authed request and read by
// bootstrap. No sheet writes; ~5 minutes of silence counts as offline.
var PRESENCE_TTL_MS = 5 * 60 * 1000;

function touchPresence_(userId) {
  try {
    var cache = CacheService.getScriptCache();
    var key = 'online_' + PROJ.id;
    var map = safeParse_(cache.get(key)) || {};
    var now = Date.now();
    map[userId] = now;
    Object.keys(map).forEach(function (id) {
      if (now - map[id] > 2 * PRESENCE_TTL_MS) delete map[id];
    });
    cache.put(key, JSON.stringify(map), 21600);
  } catch (err) { /* presence is decorative */ }
}

function onlineIds_() {
  try {
    var map = safeParse_(CacheService.getScriptCache().get('online_' + PROJ.id)) || {};
    var now = Date.now();
    return Object.keys(map).filter(function (id) { return now - map[id] < PRESENCE_TTL_MS; });
  } catch (err) {
    return [];
  }
}

function findUserByEmail_(email) {
  var users = readTable_('users');
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].email).toLowerCase() === email) return users[i];
  }
  return null;
}

/** Reverse lookup: a user row whose minted workEmail matches. Lets a person
 *  sign in with their @designthinking.lk account and still resolve to the same
 *  row they registered under with their personal email. Project-local fallback
 *  for when the cross-project directory has no row yet. */
function findUserByWorkEmail_(workEmail) {
  var key = String(workEmail || '').toLowerCase();
  if (!key) return null;
  var users = readTable_('users');
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].workEmail || '').toLowerCase() === key) return users[i];
  }
  return null;
}

/** The invites tab, created lazily — project databases predating the invite
 *  feature don't have it, and readTable_ throws on a missing tab. */
function readInvites_(noCache) {
  gid_('invites');
  return readTable_('invites', noCache);
}

function findInviteByEmail_(email) {
  var key = String(email || '').toLowerCase();
  if (!key) return null;
  var rows = readInvites_();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].email).toLowerCase() === key) return rows[i];
  }
  return null;
}

function rowById_(table, id) {
  if (!id) return null;
  var rows = readTable_(table, table === 'messages');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].id === id) return rows[i];
  }
  return null;
}

// ----------------------------------------------------------------- registry
// The central registry spreadsheet (Script Property REGISTRY_ID) is the index
// of all projects plus the cross-project people directory. Created lazily; on
// creation the pre-multi-project Script Properties (DB_ID, UPLOADS_FOLDER_ID,
// REGISTRATION_OPEN) seed the DEFAULT_PROJECT row, so an existing deployment
// migrates itself on the first request after this code ships.

function registryId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('REGISTRY_ID');
  if (id) return id;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    id = props.getProperty('REGISTRY_ID');
    if (id) return id;
    var names = Object.keys(REGISTRY_TABS);
    var ss = Sheets.Spreadsheets.create({
      properties: { title: REGISTRY_NAME },
      sheets: names.map(function (name) {
        return { properties: { title: name, gridProperties: { frozenRowCount: 1 } } };
      }),
    });
    var data = names.map(function (name) { return { range: name + '!A1', values: [REGISTRY_TABS[name]] }; });
    // Seed the default project. An existing single-project deployment donates
    // its spreadsheet/folder/config; a fresh install gets blanks (created
    // lazily on first use).
    var now = new Date().toISOString();
    data.push({ range: 'projects!A2', values: [[
      DEFAULT_PROJECT, 'ICE2026', 'Innovation & Collaboration Experience',
      'ice2026.designthinking.lk', 'active',
      getConfig_('REGISTRATION_OPEN', 'true'), 'true',
      props.getProperty('DB_ID') || '', props.getProperty('UPLOADS_FOLDER_ID') || '',
      now, now, '', '',
    ]] });
    Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'RAW', data: data }, ss.spreadsheetId);
    props.setProperty('REGISTRY_ID', ss.spreadsheetId);
    return ss.spreadsheetId;
  } finally {
    lock.releaseLock();
  }
}

/** Read a registry tab as array of objects, keyed rows only. Cached. */
function readRegistry_(tab, noCache) {
  var cache = CacheService.getScriptCache();
  if (!noCache) {
    var hit = cache.get('reg_' + tab);
    if (hit) {
      try { return JSON.parse(hit); } catch (e) { /* refetch */ }
    }
  }
  var headers = REGISTRY_TABS[tab];
  var resp = Sheets.Spreadsheets.Values.get(registryId_(), tab + '!A2:' + colLetter_(headers.length));
  var values = (resp && resp.values) || [];
  var rows = [];
  for (var r = 0; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var v = values[r][c];
      obj[headers[c]] = v === null || v === undefined ? '' : String(v);
    }
    if (obj[headers[0]]) rows.push(obj); // key column (id / email) must be set
  }
  if (!noCache) {
    var s = JSON.stringify(rows);
    if (s.length < 90000) cache.put('reg_' + tab, s, CACHE_TTL_SECONDS);
  }
  return rows;
}

function invalidateRegistry_(tab) {
  CacheService.getScriptCache().remove('reg_' + tab);
}

function appendRegistryRow_(tab, obj) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var headers = REGISTRY_TABS[tab];
    Sheets.Spreadsheets.Values.append(
      { values: [headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; })] },
      registryId_(), tab + '!A1',
      { valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS' }
    );
  } finally {
    lock.releaseLock();
  }
  invalidateRegistry_(tab);
}

/** Patch a registry row found by its key column (first header, case-insensitive). */
function updateRegistryRowByKey_(tab, keyVal, patch) {
  var found = false;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    found = updateRegistryRowUnlocked_(tab, keyVal, patch);
  } finally {
    lock.releaseLock();
  }
  invalidateRegistry_(tab);
  return found;
}

/** Lock-free inner write for callers that ALREADY hold the script lock
 *  (dbId_/uploadsFolderId_) — LockService re-entrancy is undefined. */
function updateRegistryRowUnlocked_(tab, keyVal, patch) {
  var headers = REGISTRY_TABS[tab];
  var resp = Sheets.Spreadsheets.Values.get(registryId_(), tab + '!A2:A');
  var keys = (resp && resp.values) || [];
  var rowIdx = -1;
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).toLowerCase() === String(keyVal).toLowerCase()) { rowIdx = i + 2; break; }
  }
  if (rowIdx === -1) return false;
  var range = tab + '!A' + rowIdx + ':' + colLetter_(headers.length) + rowIdx;
  var cur = ((Sheets.Spreadsheets.Values.get(registryId_(), range) || {}).values || [[]])[0] || [];
  var merged = headers.map(function (h, c) {
    return patch[h] !== undefined ? patch[h] : (cur[c] !== undefined ? cur[c] : '');
  });
  Sheets.Spreadsheets.Values.update({ values: [merged] }, registryId_(), range, { valueInputOption: 'RAW' });
  return true;
}

/** Registry row for a project slug, with booleans parsed. */
function getProject_(slug, noCache) {
  var rows = readRegistry_('projects', noCache);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].id === slug) {
      var p = rows[i];
      return {
        id: p.id,
        name: p.name,
        tagline: p.tagline,
        siteUrl: p.siteUrl,
        status: p.status || 'active',
        registrationOpen: truthy_(p.registrationOpen),
        provisionAccounts: truthy_(p.provisionAccounts),
        dbId: p.dbId,
        uploadsFolderId: p.uploadsFolderId,
        startDate: p.startDate || '',
        endDate: p.endDate || '',
      };
    }
  }
  return null;
}

/** Seed the creating admin as an admin member of the CURRENT project (PROJ must
 *  already point at the new project). Idempotent — skips if a row already
 *  exists, so a resumed create never double-seeds. Profile fields are prefilled
 *  from the cross-project directory snapshot when available. */
function seedAdminMember_(ctx) {
  if (!ctx.email) return;
  if (findUserByEmail_(ctx.email)) return;
  var now = new Date().toISOString();
  var dir = findDirectory_(ctx.email);
  var snap = {};
  try { snap = dir && dir.profile ? JSON.parse(dir.profile) : {}; } catch (e) { snap = {}; }
  appendRow_('users', {
    id: Utilities.getUuid(),
    email: ctx.email,
    name: snap.name || (dir && dir.name) || (ctx.user && ctx.user.name) || String(ctx.email).split('@')[0],
    image: snap.image || '',
    bio: snap.bio || '',
    skills: jsonArr_(snap.skills || [], 30, 40),
    affiliation: snap.affiliation || '',
    expertise: snap.expertise || '',
    gender: snap.gender || '',
    links: jsonArr_(snap.links || [], 10, 300),
    video: snap.video || '',
    role: 'admin',
    createdAt: now,
    updatedAt: now,
    workEmail: (dir && dir.workEmail) || (ctx.user && ctx.user.workEmail) || '',
  });
}

/** Create the proxied CNAME {slug}.designthinking.lk -> designthinking-lk.github.io
 *  so the central-web Worker (infra/worker.js in ice-central-web) serves the new
 *  subdomain. proxied:true is REQUIRED — Worker routes only run for proxied
 *  records. Idempotent: an already-existing record counts as success. Needs the
 *  Script Properties CLOUDFLARE_API_TOKEN + CLOUDFLARE_ZONE_ID. */
function createDnsRecord_(slug) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('CLOUDFLARE_API_TOKEN');
  var zone = props.getProperty('CLOUDFLARE_ZONE_ID');
  if (!token || !zone) {
    throw new Error('Cloudflare not configured — set CLOUDFLARE_API_TOKEN and CLOUDFLARE_ZONE_ID in Script Properties.');
  }
  var resp = UrlFetchApp.fetch(
    'https://api.cloudflare.com/client/v4/zones/' + zone + '/dns_records',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        type: 'CNAME',
        name: slug + '.designthinking.lk',
        content: 'designthinking-lk.github.io',
        proxied: true,
        comment: 'ICE project ' + slug + ' (auto-created)',
      }),
    }
  );
  var body = {};
  try { body = JSON.parse(resp.getContentText() || '{}'); } catch (e) { body = {}; }
  if (body.success) return true;
  // 81053/81057 = "record already exists" — idempotent success on resume.
  var errs = body.errors || [];
  for (var i = 0; i < errs.length; i++) {
    if (errs[i].code === 81053 || errs[i].code === 81057) return true;
  }
  var msg = errs.length ? (errs[0].code + ' ' + errs[0].message) : ('HTTP ' + resp.getResponseCode());
  throw new Error('Cloudflare DNS create failed: ' + msg);
}

/** Directory row for a personal email, or null. */
function findDirectory_(email) {
  var key = String(email || '').toLowerCase();
  if (!key) return null;
  var rows = readRegistry_('directory');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].email).toLowerCase() === key) return rows[i];
  }
  return null;
}

/** Reverse lookup: the directory row whose minted workEmail matches. This is
 *  the cross-project link between a person's personal (primary) email and their
 *  @designthinking.lk workspace account — see canonicalEmail_. */
function findDirectoryByWorkEmail_(workEmail) {
  var key = String(workEmail || '').toLowerCase();
  if (!key) return null;
  var rows = readRegistry_('directory');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].workEmail || '').toLowerCase() === key) return rows[i];
  }
  return null;
}

/** Map whatever address a person signed in with to their PRIMARY identity.
 *  Personal email (how they were invited and registered) stays primary; a
 *  minted @designthinking.lk workspace login resolves back to it via the
 *  directory. Google can't alias an external personal account to a Workspace
 *  one, so this app-level link is what makes both sign-ins the same person.
 *  Unknown addresses pass through unchanged (uninvited → invite-only card). */
function canonicalEmail_(authEmail) {
  var lower = String(authEmail || '').toLowerCase();
  if (!lower) return lower;
  var dir = findDirectoryByWorkEmail_(lower);
  return dir ? String(dir.email).toLowerCase() : lower;
}

/** Insert-or-patch a directory row. Fields not in patch are preserved. */
function upsertDirectory_(email, patch) {
  var key = String(email || '').toLowerCase();
  if (!key) return;
  patch.updatedAt = new Date().toISOString();
  if (findDirectory_(key)) {
    updateRegistryRowByKey_('directory', key, patch);
  } else {
    patch.email = key;
    appendRegistryRow_('directory', patch);
  }
}

// ------------------------------------------------------------ sheet plumbing
// All storage via the Sheets/Drive ADVANCED SERVICES so the only OAuth scope
// needed is drive.file (access limited to files this app created).

function colLetter_(n) {
  var s = '';
  while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** The current project's spreadsheet ID — created on first use, with all tabs
 *  + header rows, and written back to the project's registry row. */
function dbId_() {
  if (!PROJ) throw new Error('No project resolved — set PROJ before touching storage.');
  if (PROJ.dbId) return PROJ.dbId;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Another invocation may have created it while we waited for the lock.
    var fresh = getProject_(PROJ.id, true);
    if (fresh && fresh.dbId) { PROJ.dbId = fresh.dbId; return PROJ.dbId; }
    var names = Object.keys(TABLES);
    var ss = Sheets.Spreadsheets.create({
      properties: { title: PROJ.name + ' Database' },
      sheets: names.map(function (name) {
        return { properties: { title: name, gridProperties: { frozenRowCount: 1 } } };
      }),
    });
    var gids = {};
    (ss.sheets || []).forEach(function (sh) { gids[sh.properties.title] = sh.properties.sheetId; });
    Sheets.Spreadsheets.Values.batchUpdate({
      valueInputOption: 'RAW',
      data: names.map(function (name) { return { range: name + '!A1', values: [TABLES[name]] }; }),
    }, ss.spreadsheetId);
    PropertiesService.getScriptProperties().setProperty('DB_GIDS_' + PROJ.id, JSON.stringify(gids));
    updateRegistryRowUnlocked_('projects', PROJ.id, { dbId: ss.spreadsheetId, updatedAt: new Date().toISOString() });
    invalidateRegistry_('projects');
    PROJ.dbId = ss.spreadsheetId;
    return PROJ.dbId;
  } finally {
    lock.releaseLock();
  }
}

/** Numeric sheetId (gid) for a tab of the current project's DB; creates the
 *  tab if missing. Gid maps are cached per project in Script Properties. */
function gid_(name) {
  var props = PropertiesService.getScriptProperties();
  var key = 'DB_GIDS_' + PROJ.id;
  var raw = props.getProperty(key);
  // Pre-multi-project deployments stored the default project's map as DB_GIDS.
  if (!raw && PROJ.id === DEFAULT_PROJECT) raw = props.getProperty('DB_GIDS');
  var gids = {};
  try { gids = JSON.parse(raw || '{}'); } catch (e) { gids = {}; }
  if (gids[name] !== undefined) return gids[name];
  var meta = Sheets.Spreadsheets.get(dbId_(), { fields: 'sheets.properties' });
  gids = {};
  (meta.sheets || []).forEach(function (sh) { gids[sh.properties.title] = sh.properties.sheetId; });
  if (gids[name] === undefined) {
    var r = Sheets.Spreadsheets.batchUpdate({
      requests: [{ addSheet: { properties: { title: name, gridProperties: { frozenRowCount: 1 } } } }],
    }, dbId_());
    gids[name] = r.replies[0].addSheet.properties.sheetId;
    Sheets.Spreadsheets.Values.update({ values: [TABLES[name]] }, dbId_(), name + '!A1', { valueInputOption: 'RAW' });
  }
  props.setProperty(key, JSON.stringify(gids));
  return gids[name];
}

/** Per-project cache key for a table — projects must never share cache rows. */
function tblKey_(name) {
  return 'tbl_' + PROJ.id + '_' + name;
}

function tableRange_(name) {
  return name + '!A2:' + colLetter_(TABLES[name].length);
}

/** Read a table as array of objects. Cached unless noCache. */
function readTable_(name, noCache) {
  var cache = CacheService.getScriptCache();
  if (!noCache) {
    var hit = cache.get(tblKey_(name));
    if (hit) {
      try { return JSON.parse(hit); } catch (e) { /* refetch */ }
    }
  }
  var headers = TABLES[name];
  var resp = Sheets.Spreadsheets.Values.get(dbId_(), tableRange_(name));
  var values = (resp && resp.values) || [];
  var rows = [];
  for (var r = 0; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var v = values[r][c];
      obj[headers[c]] = v === null || v === undefined ? '' : String(v);
    }
    if (obj.id) rows.push(obj);
  }
  if (!noCache) {
    var s = JSON.stringify(rows);
    if (s.length < 90000) cache.put(tblKey_(name), s, CACHE_TTL_SECONDS);
  }
  return rows;
}

function invalidate_(name) {
  CacheService.getScriptCache().remove(tblKey_(name));
}

function appendRow_(name, obj) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var headers = TABLES[name];
    Sheets.Spreadsheets.Values.append(
      { values: [headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; })] },
      dbId_(), name + '!A1',
      { valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS' }
    );
  } finally {
    lock.releaseLock();
  }
  invalidate_(name);
}

/** 1-based sheet row index for an id (header = row 1), or -1. */
function findRowIndexById_(name, id) {
  var resp = Sheets.Spreadsheets.Values.get(dbId_(), name + '!A2:A');
  var ids = (resp && resp.values) || [];
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function updateRowById_(name, id, patch) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var headers = TABLES[name];
    var rowIdx = findRowIndexById_(name, id);
    if (rowIdx === -1) throw new Error('Row not found in ' + name + ': ' + id);
    var range = name + '!A' + rowIdx + ':' + colLetter_(headers.length) + rowIdx;
    var resp = Sheets.Spreadsheets.Values.get(dbId_(), range);
    var row = ((resp && resp.values) || [[]])[0] || [];
    var merged = headers.map(function (h, i) {
      return patch[h] !== undefined ? patch[h] : (row[i] !== undefined ? row[i] : '');
    });
    Sheets.Spreadsheets.Values.update({ values: [merged] }, dbId_(), range, { valueInputOption: 'RAW' });
  } finally {
    lock.releaseLock();
  }
  invalidate_(name);
}

function deleteRowById_(name, id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var rowIdx = findRowIndexById_(name, id);
    if (rowIdx !== -1) {
      Sheets.Spreadsheets.batchUpdate({
        requests: [{ deleteDimension: { range: {
          sheetId: gid_(name), dimension: 'ROWS', startIndex: rowIdx - 1, endIndex: rowIdx,
        } } }],
      }, dbId_());
    }
  } finally {
    lock.releaseLock();
  }
  invalidate_(name);
}

function deleteRowsWhere_(name, predicate) {
  var rows = readTable_(name, true);
  var doomed = rows.filter(predicate).map(function (r) { return r.id; });
  doomed.forEach(function (id) { deleteRowById_(name, id); });
}

/** Pull a Drive fileId out of an upload URL (profile photo / video). Returns
 *  null for external URLs (e.g. YouTube), which have no Drive file to remove. */
function driveFileId_(url) {
  var s = String(url || '');
  var m = s.match(/lh3\.googleusercontent\.com\/d\/([\w-]+)/) ||
          s.match(/drive\.google\.com\/(?:file\/d\/|uc\?[^#]*\bid=)([\w-]+)/) ||
          s.match(/\/d\/([\w-]+)/);
  return m ? m[1] : null;
}

function markMessagesRead_(ids) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var readCol = colLetter_(TABLES.messages.indexOf('read') + 1);
    var resp = Sheets.Spreadsheets.Values.get(dbId_(), 'messages!A2:A');
    var rows = (resp && resp.values) || [];
    var data = [];
    for (var i = 0; i < rows.length; i++) {
      if (ids.indexOf(String(rows[i][0])) !== -1) {
        data.push({ range: 'messages!' + readCol + (i + 2), values: [['true']] });
      }
    }
    if (data.length) {
      Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'RAW', data: data }, dbId_());
    }
  } finally {
    lock.releaseLock();
  }
}

/** Form options from the "options" tab, grouped by category. Seeds defaults
 *  on first read; rows without an id column so admins can just type values. */
function readOptions_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(tblKey_('options'));
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* refetch */ }
  }
  gid_('options'); // ensure the tab exists (creates it with the header row)
  var resp = Sheets.Spreadsheets.Values.get(dbId_(), 'options!A2:B');
  var values = (resp && resp.values) || [];
  if (!values.length) {
    var rows = [];
    Object.keys(DEFAULT_OPTIONS).forEach(function (cat) {
      DEFAULT_OPTIONS[cat].forEach(function (v) { rows.push([cat, v]); });
    });
    Sheets.Spreadsheets.Values.append({ values: rows }, dbId_(), 'options!A1',
      { valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS' });
    values = rows;
  }
  var out = {};
  values.forEach(function (r) {
    var cat = String(r[0] || '').trim().toLowerCase();
    var val = String((r[1] === undefined ? '' : r[1])).trim();
    if (!cat || !val) return;
    if (!out[cat]) out[cat] = [];
    if (out[cat].indexOf(val) === -1) out[cat].push(val);
  });
  cache.put(tblKey_('options'), JSON.stringify(out), CACHE_TTL_SECONDS);
  return out;
}

/** The six #/projects cards, seeded from DEFAULT_TEAM_PROJECTS on first read. */
function readTeamProjects_() {
  gid_('team_projects'); // ensure the tab exists (creates it with the header row)
  var rows = readTable_('team_projects', true);
  if (!rows.length) {
    var now = new Date().toISOString();
    DEFAULT_TEAM_PROJECTS.forEach(function (p, i) {
      appendRow_('team_projects', {
        id: Utilities.getUuid(), slot: String(i),
        title: p.title, description: p.description, color: 'pc-' + (i + 1),
        updatedBy: '', updatedAt: now, video: '', fullDescription: '', website: '', websiteOk: '',
      });
    });
    rows = readTable_('team_projects', true);
  }
  return rows.map(function (p) {
    var out = {};
    Object.keys(p).forEach(function (k) { out[k] = p[k]; });
    out.slot = Number(p.slot) || 0;
    return out;
  }).sort(function (a, b) { return a.slot - b.slot; });
}

/** True if a URL responds without a 404/410 (or a hard connection failure).
 *  Same policy as the check_url action — bot-blocked hosts still count as live. */
function urlReachable_(url) {
  if (!/^https?:\/\//i.test(url)) return false;
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'get', followRedirects: true, muteHttpExceptions: true, validateHttpsCertificates: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; ICE-linkcheck/1.0; +https://ice.designthinking.lk)' },
    });
    var code = resp.getResponseCode();
    return !(code === 404 || code === 410);
  } catch (err) {
    return false;
  }
}

/** The team that owns project slot i: teams sorted by name (== the frontend's
 *  homeTeams()) taken by index. undefined when fewer teams than slots exist. */
function teamForSlot_(slot) {
  var teams = readTable_('teams').slice().sort(function (a, b) {
    return String(a.name || '').localeCompare(String(b.name || ''));
  });
  return teams[slot];
}

// Community roles let someone sit in a team; admin-only people can't. Used to
// decide auto-unassignment when a role is removed.
function isCommunityMember_(user) {
  var roles = rolesOf_(user);
  return roles.indexOf('participant') !== -1 || roles.indexOf('mentor') !== -1;
}

// Pull a user out of every team they're on (they lost their last community
// role). Returns how many teams were touched.
function removeUserFromTeams_(uid) {
  var now = new Date().toISOString();
  var removed = 0;
  readTable_('teams', true).forEach(function (t) {
    var members = parseArr_(t.members);
    if (members.indexOf(uid) !== -1) {
      updateRowById_('teams', t.id, {
        members: JSON.stringify(members.filter(function (id) { return id !== uid; })),
        updatedAt: now,
      });
      removed++;
    }
  });
  return removed;
}

// ----------------------------------------------------------- persona (LLM)
// Claude writes the short persona blurb shown beside the card while a person
// fills in their profile. Raw Messages API over UrlFetchApp (no Apps Script
// SDK exists). Key: Script Property ANTHROPIC_API_KEY. Swap PERSONA_MODEL to
// 'claude-haiku-4-5' if per-keystroke cost ever matters more than quality.
var PERSONA_MODEL = 'claude-opus-4-8';
// Skill blurbs are tiny and cached hard — haiku is plenty.
var SKILL_MODEL = 'claude-haiku-4-5';

function generateSkillBlurb_(apiKey, skill) {
  try {
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: SKILL_MODEL,
        max_tokens: 130,
        system: 'You explain skills to participants of a design-thinking innovation workshop. In one or two plain, friendly sentences (at most ~35 words), say what the given skill is and why it helps when building a project. Respond with only the description — no preamble, no quotes, no markdown.',
        messages: [{ role: 'user', content: 'Skill: ' + skill }],
      }),
    });
    var code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      console.error('skill blurb HTTP ' + code + ': ' + resp.getContentText().slice(0, 300));
      return '';
    }
    var data = JSON.parse(resp.getContentText());
    if (data.stop_reason === 'refusal') return '';
    var out = '';
    (data.content || []).forEach(function (b) { if (b.type === 'text') out += b.text; });
    return out.trim();
  } catch (err) {
    console.error('generateSkillBlurb_ failed: ' + ((err && err.stack) || err));
    return '';
  }
}

function generatePersona_(apiKey, fields) {
  try {
    var lines = [];
    if (fields.name) lines.push('Name: ' + fields.name);
    lines.push('Role at the workshop: ' + fields.role);
    if (fields.affiliation) lines.push('Affiliation: ' + fields.affiliation);
    if (fields.expertise) lines.push('Expertise: ' + fields.expertise);
    if (fields.skills.length) lines.push('Skills: ' + fields.skills.join(', '));
    if (fields.bio) lines.push('Bio: ' + fields.bio);
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: PERSONA_MODEL,
        max_tokens: 300,
        system: 'You write short persona introductions for a design-thinking workshop\'s community platform. From the profile fields provided, write a warm, positive, third-person introduction of this person in 2-3 sentences (at most ~60 words). Celebrate what is there; never mention missing or empty fields, never invent facts. If the fields are very sparse, write one inviting sentence about who they seem to be so far. Respond with only the introduction text - no preamble, no quotes, no markdown.',
        messages: [{ role: 'user', content: lines.join('\n') }],
      }),
    });
    var code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      console.error('persona API HTTP ' + code + ': ' + resp.getContentText().slice(0, 300));
      return '';
    }
    var data = JSON.parse(resp.getContentText());
    if (data.stop_reason === 'refusal') return '';
    var out = '';
    (data.content || []).forEach(function (b) { if (b.type === 'text') out += b.text; });
    return out.trim();
  } catch (err) {
    console.error('generatePersona_ failed: ' + ((err && err.stack) || err));
    return '';
  }
}

// ------------------------------------------------- workspace provisioning
// Creates a Google Workspace account firstname@designthinking.lk in the /ICE
// org unit via the Admin SDK Directory advanced service (AdminDirectory), then
// emails the temporary password to the address the person signed in with.
// Requires: the api project's owner is a Workspace super-admin, designthinking.lk
// is a verified domain, and the /ICE org unit exists. Scopes: admin.directory.user
// + script.send_mail (see appsscript.json). Always returns '' on any failure so a
// registration is never blocked by provisioning problems.

/** Reduce a name part to a bare email-handle token: lowercase, [a-z0-9] only. */
function handlePart_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function isDuplicateUserError_(err) {
  var m = String((err && err.message) || err || '');
  return /already exist|duplicate|entity.*exist|409/i.test(m);
}

/** 16-char password satisfying default Workspace complexity (letter+digit+symbol). */
function randomPassword_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  var s = '';
  for (var i = 0; i < 14; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s + 'q7$';
}

/** Create the workshop account, trying firstname@ then firstname.lastname@ then
 *  numbered variants on collision. Returns the created email, or '' on failure. */
function provisionWorkspaceAccount_(first, last, notifyEmail) {
  try {
    if (typeof AdminDirectory === 'undefined') return '';
    var f = handlePart_(first);
    var l = handlePart_(last);
    if (!f) return '';
    var candidates = [f];
    if (l) {
      candidates.push(f + '.' + l);
      for (var n = 2; n <= 20; n++) candidates.push(f + '.' + l + n);
    } else {
      for (var n2 = 2; n2 <= 20; n2++) candidates.push(f + n2);
    }
    var password = randomPassword_();
    for (var i = 0; i < candidates.length; i++) {
      var primaryEmail = candidates[i] + '@' + WORKSPACE_DOMAIN;
      try {
        AdminDirectory.Users.insert({
          primaryEmail: primaryEmail,
          name: { givenName: first || f, familyName: last || first || f },
          password: password,
          changePasswordAtNextLogin: true,
          orgUnitPath: WORKSPACE_OU,
        });
        sendWorkspaceCreds_(notifyEmail, first || f, primaryEmail, password);
        return primaryEmail;
      } catch (err) {
        if (isDuplicateUserError_(err)) continue; // handle taken — try the next one
        throw err; // real error (auth/scope/domain) — abort, caught below
      }
    }
    return '';
  } catch (err) {
    console.error('provisionWorkspaceAccount_ failed: ' + ((err && err.stack) || err));
    return '';
  }
}

/** Email the new workshop credentials to the address the person signed in with. */
function sendWorkspaceCreds_(to, firstName, workEmail, password) {
  if (!to) return;
  try {
    var ev = escapeHtmlA_(PROJ.name);
    var html =
      '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:0 auto;color:#0E0F11">' +
      '<h2 style="color:#6100FF;margin:0 0 6px">Your ' + ev + ' workshop account</h2>' +
      '<p>Hi ' + escapeHtmlA_(firstName) + ',</p>' +
      '<p>We’ve created a workshop Google account for you as part of your ' + ev + ' registration. Use it to sign in to the ' + ev + ' site and to message mentors and other participants in Google Chat.</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:18px 0;border-collapse:collapse">' +
      '<tr><td style="padding:8px 14px;background:#F4F1FB;border-radius:8px 8px 0 0;font-size:13px;color:#555">Sign in with</td></tr>' +
      '<tr><td style="padding:12px 14px;background:#F8F7FC;font-size:16px"><b>' + escapeHtmlA_(workEmail) + '</b></td></tr>' +
      '<tr><td style="padding:12px 14px;background:#F4F1FB;border-radius:0 0 8px 8px;font-size:16px">Temporary password: <b>' + escapeHtmlA_(password) + '</b></td></tr>' +
      '</table>' +
      '<p style="font-size:14px;color:#555">You’ll be asked to set a new password on first sign-in. It’s the same registration either way — signing in to the ' + ev + ' site with this account or with your own Google account (' + escapeHtmlA_(to) + ') both bring you to your profile.</p>' +
      '<p style="font-size:13px;color:#888;margin-top:22px">' + ev + ' · Augmented Human Lab</p>' +
      '</div>';
    MailApp.sendEmail({
      to: to,
      subject: 'Your ' + PROJ.name + ' workshop account',
      htmlBody: html,
      name: PROJ.name,
    });
  } catch (err) {
    console.error('sendWorkspaceCreds_ failed: ' + ((err && err.stack) || err));
  }
}

/** Onboarding invitation: sign in with THIS email on the workshop site, then
 *  complete the profile card. Returns false when the send fails (quota etc.)
 *  so the caller can report it — the allowlist row is written regardless.
 *  Deliverability: a plain-text body ships alongside the HTML and the link is
 *  also visible as text (HTML-only, button-only mail scores high on spam
 *  filters); replies go to the organizer who sent the invite. */
function sendInviteEmail_(to, role, replyTo) {
  try {
    var ev = escapeHtmlA_(PROJ.name);
    var url = PROJ.siteUrl
      ? (/^https?:\/\//i.test(PROJ.siteUrl) ? PROJ.siteUrl : 'https://' + PROJ.siteUrl)
      : 'https://ice.designthinking.lk/?project=' + PROJ.id;
    var roleLabel = role === 'mentor' ? 'a mentor' : role === 'admin' ? 'an organizer' : 'a participant';
    var html =
      '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:0 auto;color:#0E0F11">' +
      '<h2 style="color:#6100FF;margin:0 0 6px">You&#39;re invited to ' + ev + '</h2>' +
      '<p>Hello,</p>' +
      '<p>The organizers have invited you to join <b>' + ev + '</b> as <b>' + roleLabel + '</b>. Complete your registration to meet the other ' +
      (role === 'mentor' ? 'mentors and participants' : 'participants and mentors') + ' and get started.</p>' +
      '<ol style="font-size:14.5px;line-height:1.7;padding-left:20px;margin:16px 0">' +
      '<li>Open the workshop site.</li>' +
      '<li>Sign in with Google using <b>this email address</b> (' + escapeHtmlA_(to) + ') — only invited addresses can register.</li>' +
      '<li>Fill in your profile card and join.</li>' +
      '</ol>' +
      '<p style="margin:22px 0 8px"><a href="' + escapeHtmlA_(url) + '" style="display:inline-block;padding:12px 30px;border-radius:999px;font-size:14.5px;font-weight:600;color:#ffffff;text-decoration:none;background:#6100FF">Join ' + ev + '</a></p>' +
      '<p style="font-size:13px;color:#555;margin:0 0 8px">Or open this link: <a href="' + escapeHtmlA_(url) + '">' + escapeHtmlA_(url) + '</a></p>' +
      '<p style="font-size:13px;color:#888;margin-top:22px">' + ev + ' · Augmented Human Lab</p>' +
      '</div>';
    var text =
      'You’re invited to ' + PROJ.name + '\n\n' +
      'The organizers have invited you to join ' + PROJ.name + ' as ' + roleLabel + '.\n\n' +
      '1. Open the workshop site: ' + url + '\n' +
      '2. Sign in with Google using this email address (' + to + ') — only invited addresses can register.\n' +
      '3. Fill in your profile card and join.\n\n' +
      PROJ.name + ' · Augmented Human Lab';
    var msg = {
      to: to,
      // Unique per send (time-stamped) so Gmail doesn't thread repeat invites
      // to the same address into one collapsed conversation.
      subject: 'You’re invited to ' + PROJ.name + ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a') + ')',
      body: text,
      htmlBody: html,
      name: PROJ.name,
    };
    if (replyTo && EMAIL_RE.test(replyTo)) msg.replyTo = replyTo;
    MailApp.sendEmail(msg);
    return true;
  } catch (err) {
    console.error('sendInviteEmail_ failed: ' + ((err && err.stack) || err));
    return false;
  }
}

/** Returning person: they already have a @designthinking.lk account from an
 *  earlier workshop — remind them it works here too, instead of minting a
 *  duplicate. No password included; they keep their existing one. */
function sendWorkspaceWelcomeBack_(to, firstName, workEmail) {
  if (!to) return;
  try {
    var ev = escapeHtmlA_(PROJ.name);
    var html =
      '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:0 auto;color:#0E0F11">' +
      '<h2 style="color:#6100FF;margin:0 0 6px">Welcome back to ' + ev + '</h2>' +
      '<p>Hi ' + escapeHtmlA_(firstName) + ',</p>' +
      '<p>Good news — the workshop account you got at a previous workshop works for ' + ev + ' too. Use it to sign in to the ' + ev + ' site and to message people in Google Chat.</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:18px 0;border-collapse:collapse">' +
      '<tr><td style="padding:8px 14px;background:#F4F1FB;border-radius:8px 8px 0 0;font-size:13px;color:#555">Sign in with</td></tr>' +
      '<tr><td style="padding:12px 14px;background:#F8F7FC;border-radius:0 0 8px 8px;font-size:16px"><b>' + escapeHtmlA_(workEmail) + '</b></td></tr>' +
      '</table>' +
      '<p style="font-size:14px;color:#555">Use the password you set last time — or just sign in with your own Google account (' + escapeHtmlA_(to) + '); both reach the same ' + ev + ' registration. Forgotten the password? Reply to this email and the organizers will reset it for you.</p>' +
      '<p style="font-size:13px;color:#888;margin-top:22px">' + ev + ' · Augmented Human Lab</p>' +
      '</div>';
    MailApp.sendEmail({
      to: to,
      subject: 'Your ' + PROJ.name + ' workshop account',
      htmlBody: html,
      name: PROJ.name,
    });
  } catch (err) {
    console.error('sendWorkspaceWelcomeBack_ failed: ' + ((err && err.stack) || err));
  }
}

function escapeHtmlA_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function uploadsFolderId_() {
  if (PROJ.uploadsFolderId) return PROJ.uploadsFolderId;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var fresh = getProject_(PROJ.id, true);
    if (fresh && fresh.uploadsFolderId) { PROJ.uploadsFolderId = fresh.uploadsFolderId; return PROJ.uploadsFolderId; }
    var folder = Drive.Files.create({ name: PROJ.name + ' Uploads', mimeType: 'application/vnd.google-apps.folder' });
    updateRegistryRowUnlocked_('projects', PROJ.id, { uploadsFolderId: folder.id, updatedAt: new Date().toISOString() });
    invalidateRegistry_('projects');
    PROJ.uploadsFolderId = folder.id;
    return PROJ.uploadsFolderId;
  } finally {
    lock.releaseLock();
  }
}

/** Run once from the IDE to authorize all scopes (drive.file, admin.directory.user,
 *  send_mail), create the registry (migrating an existing single-project
 *  deployment into it) and the default project's database. Re-run after adding
 *  scopes so Google shows the consent screen for the new permissions. */
function setup() {
  console.log('Registry ready: https://docs.google.com/spreadsheets/d/' + registryId_());
  PROJ = getProject_(DEFAULT_PROJECT, true);
  if (!PROJ) throw new Error('Default project missing from registry: ' + DEFAULT_PROJECT);
  var id = dbId_();
  console.log('Database ready: https://docs.google.com/spreadsheets/d/' + id);
  console.log('Uploads folder id: ' + uploadsFolderId_());
  console.log('Workspace check: ' + checkWorkspaceAccess());
}

/** Re-write the registry tabs' header rows — run from the IDE after
 *  REGISTRY_TABS gains new columns (data alignment is unaffected because new
 *  columns are always appended at the end). */
function patchRegistryHeaders() {
  var id = registryId_();
  Sheets.Spreadsheets.Values.batchUpdate({
    valueInputOption: 'RAW',
    data: Object.keys(REGISTRY_TABS).map(function (name) {
      return { range: name + '!A1', values: [REGISTRY_TABS[name]] };
    }),
  }, id);
  console.log('Registry headers updated.');
}

/** One-shot, idempotent: backfill the registry's cross-project directory from
 *  the default project's existing users tab. Run from the IDE after the
 *  multi-project code first ships. Rows already in the directory are left
 *  untouched so a re-run never clobbers newer data. */
function migrateDirectoryFromUsers() {
  registryId_();
  PROJ = getProject_(DEFAULT_PROJECT, true);
  if (!PROJ || !PROJ.dbId) throw new Error('No ' + DEFAULT_PROJECT + ' database to migrate from.');
  var added = 0;
  readTable_('users', true).forEach(function (u) {
    if (!u.email || findDirectory_(u.email)) return;
    upsertDirectory_(u.email, {
      workEmail: u.workEmail || '',
      name: u.name,
      lastProjectId: PROJ.id,
      profile: JSON.stringify(profileSnapshot_(u)),
    });
    added++;
  });
  console.log('Directory backfilled: ' + added + ' added, ' + readRegistry_('directory', true).length + ' total.');
}

/** Smoke-test the Admin SDK wiring without creating anyone. Reads one account in
 *  the workshop domain — this validates super-admin directory access and that
 *  designthinking.lk is a domain in this Workspace, using only the
 *  admin.directory.user scope that Users.insert also needs (no extra scope). The
 *  /ICE org unit is exercised for real at insert time (provisioning is guarded). */
function checkWorkspaceAccess() {
  try {
    if (typeof AdminDirectory === 'undefined') return 'AdminDirectory advanced service is NOT enabled.';
    var resp = AdminDirectory.Users.list({ customer: 'my_customer', domain: WORKSPACE_DOMAIN, maxResults: 1 });
    var n = (resp && resp.users && resp.users.length) || 0;
    return 'OK — ' + WORKSPACE_DOMAIN + ' reachable (' + n + ' account' + (n === 1 ? '' : 's') + ' visible); ready to mint accounts into ' + WORKSPACE_OU + '.';
  } catch (err) {
    return 'FAILED — ' + ((err && err.message) || err) + ' (check super-admin rights and that ' + WORKSPACE_DOMAIN + ' is a verified domain).';
  }
}
