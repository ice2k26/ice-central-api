# ICE — Central API (`ice-central-api`, Google Apps Script)

Two clasp-managed Apps Script web apps, deployed from sankha@ahlab.org:

| Project | Executes as | Access | Purpose |
|---|---|---|---|
| `auth/` | User accessing | Anyone with Google account | Google sign-in broker — mints HMAC-signed tokens, redirects back to the site with `#icetoken=` |
| `api/`  | Owner | Anyone (anonymous) | Multi-project JSON API over Google Sheets (users, teams, messages, announcements) + Drive image uploads |

The two projects share an HMAC secret (generate with `openssl rand -hex 32`):
in `api/` it lives in `src/Secret.js` (git-ignored — copy `Secret.js.example`);
in `auth/` it lives in a **Script Property** named `SECRET`, never in source.

**Why the split, and why `auth/` must be shared:** the auth web app executes
as *user accessing*, and Apps Script only lets a visitor run such a web app
if they have at least **view access to the script file itself** — otherwise
they get "Sorry, unable to open the file at present". The auth script is
therefore shared "anyone with the link — viewer", which makes its source
world-readable; the secret sits in Script Properties, which viewers can't
read. (The api project executes as owner, needs no sharing, so its
`Secret.js` stays private.)

## Multi-project storage

A central **registry spreadsheet** (Script Property `REGISTRY_ID`, created
lazily) indexes every project (workshop instance) and holds a cross-project
people directory:

- `projects` tab — one row per project: slug, name, tagline, siteUrl,
  status (`active`/`test`/`archived`), registrationOpen, provisionAccounts,
  and pointers to the project's own database spreadsheet + uploads folder
  (both auto-created on first use).
- `directory` tab — one row per person (keyed by the personal email they
  sign in with): their minted `@designthinking.lk` account and a profile
  snapshot. Returning registrants keep their existing account (no duplicate
  mint) and get their form prefilled.

Every API request carries `project: '<slug>'`; omitting it falls back to
`ice2026`. Test projects set `provisionAccounts` off, so registering there
never creates real Workspace accounts or sends credential emails.

## Workflows

```bash
cd api   # or auth
npm install
npx clasp push --force   # upload code
node deploy.js           # push + redeploy to the fixed deployment ID (stable /exec URL)
```

CORS: the frontend POSTs with `Content-Type: text/plain` (a CORS "simple request"),
which Apps Script answers with `Access-Control-Allow-Origin: *`. No preflight, no proxy.

One-time setup after creating from scratch: run `setup()` in the api project's
script editor to authorize scopes, create the registry (seeding the default
project from any pre-multi-project `DB_ID`) and the default project's
database. After the multi-project code first ships over an existing
deployment, also run `migrateDirectoryFromUsers()` once to backfill the
directory from the existing users tab.
