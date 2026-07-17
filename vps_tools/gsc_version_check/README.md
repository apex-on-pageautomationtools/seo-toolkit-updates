# GSC Version Check (VPS)

Moves the GSC "version check" sweep (www/non-www, http/https correctness -
originally `_runVersionCheckSweep` in the Apps Script) off Apps Script and
onto your VPS, since it was hitting Apps Script's daily UrlFetchApp quota
running against a large `GSC_Config` sheet. The Apps Script version stays in
place as a fallback/manual option - this becomes the primary path.

**Nothing existing is touched.** No new Google Sheet is created, no existing
OAuth token is removed or invalidated. This uses:
- A **copy** of an already-connected GSC account's refresh token, just to
  call the Search Console API from here - reusing a refresh token from a
  second place doesn't revoke or affect its original use.
- A **new, separate** Google service account, used ONLY to read/write the
  same existing `GSC_Config` sheet (shared with it as an Editor, same as
  sharing a sheet with any other collaborator) - it never touches Search
  Console at all.

## One-time setup

### 1. Get the GSC OAuth values (copy, don't move)

On the machine currently running SEO Toolkit Pro:
- `gsc_client_id` / `gsc_client_secret`: in `config.json` at the repo root
  (same values already used for every other GSC OAuth flow in the app).
- `gsc_refresh_token`: open `.gsc_accounts` at the repo root, find the
  account you want this script to check with, copy its `refresh_token`
  value. **This is a copy** - the original keeps working exactly as before
  in the main app.

### 2. Create the service account (Google Cloud Console)

1. In the same Google Cloud project the GSC OAuth client already lives in
   (or a new one - doesn't matter, service accounts are independent):
   APIs & Services -> Credentials -> Create Credentials -> Service Account.
2. Give it any name (e.g. "gsc-version-check-sheets"). No roles needed at
   the project level.
3. Open the new service account -> Keys -> Add Key -> Create new key -> JSON.
   Download it, copy it to the VPS (e.g. `/opt/gsc_version_check/service_account.json`).
4. Enable the **Google Sheets API** for this project if not already enabled
   (APIs & Services -> Library -> search "Google Sheets API" -> Enable).
5. Open the `GSC_Config` Google Sheet in your browser -> Share -> paste the
   service account's email (looks like
   `gsc-version-check-sheets@your-project.iam.gserviceaccount.com`, also in
   the downloaded JSON as `client_email`) -> give it **Editor** access.

### 3. Deploy to the VPS

```bash
mkdir -p /opt/gsc_version_check
# copy gsc_version_check.py and config.example.json there
cd /opt/gsc_version_check
python3 -m venv venv
source venv/bin/activate
pip install cryptography
cp config.example.json config.json
nano config.json   # fill in sheet_id, service_account_key_file path, gsc_client_id/secret/refresh_token
```

`sheet_id` is the long ID in the sheet's URL:
`https://docs.google.com/spreadsheets/d/`**`THIS_PART`**`/edit`

### 4. Test it once, manually

```bash
cd /opt/gsc_version_check
source venv/bin/activate
python3 gsc_version_check.py
```

Watch the log output - it should report how many domains needed a check and
how many it processed. Check the `GSC_Config` sheet afterward to confirm
`Version_Check_Status`/`Version_Check_At`/`Version_Check_Detail` got filled
in for a batch of rows.

### 5. Schedule it (cron)

```bash
crontab -e
```

Add a line to run it every 15 minutes (adjust to taste - each run only
processes up to 200 domains at a time, so a large sheet sweeps over several
runs rather than one long one):

```
*/15 * * * * cd /opt/gsc_version_check && venv/bin/python3 gsc_version_check.py >> run.log 2>&1
```

## How it decides Correct vs Wrong Version

Identical logic to the Apps Script's `_prepareDomainForVersionCheck` /
`_resultFromInspection`:

1. If the connected account has a **domain property** (`sc-domain:...`) for
   this domain -> **Correct** (a domain property covers every URL variant,
   no live check needed).
2. If the account has **no property at all** for this domain -> **Wrong
   Version**.
3. If the account has a **URL-prefix property** that might not match the
   live site's actual URL - call the URL Inspection API and compare its
   GSC-verified canonical URL's host against the registered property's
   host. Match -> **Correct**. Mismatch -> **Wrong Version**.
4. Anything inconclusive (API error, not yet crawled by Google) is left
   blank and retried on the next run - never guessed.

## Only checks domains with a blank Version_Check_Status

Same as the Apps Script's targeted (not full-sheet) sweep - it only
processes rows where `Version_Check_Status` is empty. To force a full
re-check of every domain, clear that column in the sheet first (or add a
`--force` flag yourself if you want that - not built in, to avoid an
accidental full-sheet API burst).
