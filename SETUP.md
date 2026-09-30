# Setup and operations

## Offline first

Follow the README to install dependencies, run `python demo.py`, and run `python -m unittest discover -v`. Use Python 3.12 or 3.13. No Google credentials are needed for those steps.

Generate a fictional enrollment workbook with `python demo.py --write-sample`. Its columns are `Section`, `CourseCode`, `StudentId` and `Name`. Replace it with your own authorized workbook for a real deployment; keep real rosters outside version control.

## Connect your own Google project

Enable Calendar, Sheets and Drive APIs. Configure an OAuth web application and add `http://localhost:5050/oauth2callback` as an allowed redirect URI. While in OAuth testing mode, add your test account to the consent screen. Store the downloaded OAuth client JSON outside the repository.

The current scopes include email, Calendar event writes, Sheets reads and Drive reads. Review whether that access is appropriate for your own deployment. Never upload access tokens, refresh tokens, OAuth client secrets, real student rosters or personal calendars.

```bash
python demo.py --write-sample
export GOOGLE_SHEET_ID="your-own-timetable-file-id"
export GOOGLE_CLIENT_SECRET_FILE="/absolute/path/outside-repo/client_secret.json"
export SECTION_WORKBOOK="$PWD/sample_sections.xlsx"
export FLASK_SECRET_KEY="$(python -c 'import secrets; print(secrets.token_urlsafe(48))')"
export OAUTHLIB_INSECURE_TRANSPORT=1
export TZ=Asia/Kolkata
python app.py
```

Open `http://localhost:5050`. `OAUTHLIB_INSECURE_TRANSPORT=1` is for localhost development only. Do not enable it in a hosted environment. The `.env.example` file documents configuration; the application does not automatically load `.env` files.

## Timetable format

The current parser expects week tabs such as `Week (11-17 Jan)`, a header row with time ranges such as `09:00 AM - 10:30 AM`, day names in the first column, and course cells such as:

```text
26MBA401: Digital Transformation [A]
Prof. Example
(Demo Room 1)
```

It reads `A1:H40` from selected current or future weekly tabs. Strikethrough cancellation formatting is read through the Drive XLSX fallback; the ordinary Sheets values path does not retrieve formatting. Confirm which path your source uses before relying on cancellation handling.

## Container

```bash
docker build -t campus-calendar-sync .
docker run --rm -p 8080:8080 \
  -e FLASK_SECRET_KEY="your-long-random-secret" \
  -e TZ=Asia/Kolkata \
  campus-calendar-sync
```

This starts the UI and health endpoint with a synthetic enrollment workbook. Real sync additionally needs your Google configuration and a mounted OAuth secret or `GOOGLE_CLIENT_SECRET_JSON` injected at runtime.

For Cloud Run, use Secret Manager for `FLASK_SECRET_KEY` and `GOOGLE_CLIENT_SECRET_JSON`, enable secure cookies, and register the hosted `/oauth2callback` redirect URI. The code supports `OAUTH_REDIRECT_URI`, `SESSION_COOKIE_SECURE`, `TIMETABLE_CACHE_TTL_SECONDS`, `MAX_DAILY_SYNC_ATTEMPTS` and `SYNC_TRACKER_DB`.

Keep one instance while using the local SQLite daily counter; multiple instances require a shared store. The five-minute timetable cache and in-process locks are per container. `/healthz` checks process health, not Google API access.

## Known security and data limitations

Use a strong secret key: the development fallback is not suitable for deployment. OAuth credentials are encrypted inside the signed browser session; a server-side credential store is a future improvement. Enrollment lookup by name or ID does not establish that the signed-in account owns that enrollment. Broader release should include identity binding, a privacy review, appropriate OAuth consent and a clear revocation flow.
