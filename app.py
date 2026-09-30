import hashlib
import hmac
import json
import logging
import os
import base64
import sys
from dataclasses import asdict
from pathlib import Path
from tempfile import NamedTemporaryFile
from threading import Lock
from urllib import request as url_request
from urllib.parse import urlencode

from cryptography.fernet import Fernet, InvalidToken
from flask import Flask, jsonify, redirect, render_template_string, request, session, url_for
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from oauthlib.oauth2.rfc6749.errors import MismatchingStateError
from werkzeug.middleware.proxy_fix import ProxyFix

from services import CalendarSyncService, SectionRepository, TimetableService
from sync_tracker import check_and_increment_sync, init_db


BASE_DIR = Path(__file__).resolve().parent
SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
WORKBOOK_PATH = os.environ.get(
    "SECTION_WORKBOOK",
    str(BASE_DIR / "sample_sections.xlsx"),
)
CLIENT_SECRET_FILE = os.environ.get("GOOGLE_CLIENT_SECRET_FILE", "")
CLIENT_SECRET_JSON = os.environ.get("GOOGLE_CLIENT_SECRET_JSON", "")
OAUTH_REDIRECT_URI = os.environ.get("OAUTH_REDIRECT_URI", "")
SUPPORT_UPI_ID = os.environ.get("SUPPORT_UPI_ID", "").strip()
SUPPORT_PAYEE_NAME = os.environ.get("SUPPORT_PAYEE_NAME", "Calendar Sync").strip()
SUPPORT_NOTE = os.environ.get("SUPPORT_NOTE", "Coffee for timetable sync app").strip()
ENABLE_SUPPORT_PREVIEW = os.environ.get("ENABLE_SUPPORT_PREVIEW", "0") == "1"
MAX_DAILY_SYNC_ATTEMPTS = int(os.environ.get("MAX_DAILY_SYNC_ATTEMPTS", "5"))
SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "local-dev-secret-change-me")
if not app.logger.handlers:
    app.logger.addHandler(logging.StreamHandler(sys.stdout))
app.logger.setLevel(logging.INFO)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("SESSION_COOKIE_SECURE", "0") == "1",
    PREFERRED_URL_SCHEME="https",
)
_client_secret_tempfile = None
_sync_locks: dict[str, Lock] = {}
_sync_locks_guard = Lock()
init_db()


PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>IIM Calendar Sync</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #17202a;
      --muted: #64717f;
      --line: #dbe3eb;
      --soft-line: #edf2f6;
      --surface: #ffffff;
      --page: #f5f7f9;
      --green: #0b8043;
      --green-dark: #075b32;
      --green-soft: #e7f4ed;
      --blue-soft: #eaf2ff;
      --blue: #2463eb;
      --warn: #915900;
      --warn-soft: #fff5db;
      --error: #a62626;
      --error-soft: #fde8e8;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    * { box-sizing: border-box; }
    body { margin: 0; background: #ffffff; color: var(--ink); }
    .site-nav {
      max-width: 1160px;
      margin: 0 auto;
      padding: 28px 20px 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 24px;
    }
    .brand-logo { width: 154px; height: auto; display: block; }
    .nav-actions { display: flex; align-items: center; justify-content: flex-end; gap: 12px; }
    .hero {
      min-height: 210px;
      display: grid;
      place-items: center;
      background:
        linear-gradient(180deg, rgba(247, 252, 255, .78), rgba(224, 242, 255, .85)),
        linear-gradient(120deg, #e7f4ed, #eaf2ff);
      border-top: 1px solid #eef3f6;
      border-bottom: 1px solid #dcebf5;
      text-align: center;
      padding: 34px 20px;
    }
    .hero h1 {
      margin: 0;
      max-width: 980px;
      color: #082d38;
      font-family: Georgia, "Times New Roman", serif;
      font-size: clamp(42px, 7vw, 82px);
      line-height: .98;
      letter-spacing: 0;
      font-weight: 800;
      text-transform: capitalize;
    }
    main { max-width: 1160px; margin: 0 auto; padding: 32px 20px 56px; position: relative; }
    h2 { margin: 0; font-size: 18px; letter-spacing: 0; }
    h3 { margin: 0; font-size: 14px; letter-spacing: 0; }
    .muted { color: var(--muted); font-size: 14px; line-height: 1.5; }
    .layout { max-width: 920px; margin: 0 auto; display: grid; gap: 18px; }
    .panel { background: rgba(255,255,255,.96); border: 1px solid var(--line); border-radius: 8px; box-shadow: 0 1px 2px rgba(18, 31, 45, .04); }
    .panel-pad { padding: 22px; }
    .sync-card { display: grid; gap: 18px; }
    .title-row { display: flex; justify-content: space-between; gap: 14px; align-items: start; }
    .eyebrow { color: var(--green-dark); font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: .08em; margin-bottom: 6px; }
    .status-chip { display: inline-flex; align-items: center; gap: 8px; height: 30px; padding: 0 10px; border-radius: 999px; background: var(--green-soft); color: var(--green-dark); font-size: 12px; font-weight: 800; white-space: nowrap; }
    .dot { width: 8px; height: 8px; border-radius: 99px; background: var(--green); }
    .form-grid { display: grid; grid-template-columns: minmax(0, 1fr) auto auto; gap: 10px; align-items: end; }
    label { display: block; font-size: 13px; font-weight: 750; margin-bottom: 7px; }
    input {
      width: 100%;
      border: 1px solid #bcc8d4;
      border-radius: 7px;
      padding: 12px 13px;
      font-size: 15px;
      line-height: 1.2;
      background: #fff;
      color: var(--ink);
    }
    input:focus { outline: 3px solid rgba(11, 128, 67, .16); border-color: var(--green); }
    button, a.button {
      min-height: 43px;
      border: 0;
      background: var(--green);
      color: white;
      border-radius: 7px;
      padding: 0 14px;
      font-weight: 800;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      font-size: 14px;
      white-space: nowrap;
    }
    button:hover, a.button:hover { background: var(--green-dark); }
    button.secondary { background: #e8eef4; color: #233241; }
    button.secondary:hover { background: #dbe4ed; }
    button.danger { background: var(--error); color: #fff; }
    button.danger:hover { background: #7f1d1d; }
    button:disabled { opacity: .65; cursor: not-allowed; }
    .status { min-height: 24px; font-weight: 750; font-size: 14px; }
    .status.ok { color: var(--green-dark); }
    .status.error { color: var(--error); }
    .status.loading { color: var(--blue); }
    .trust-row { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; }
    .trust-item { border: 1px solid var(--soft-line); border-radius: 8px; padding: 12px; background: #fbfcfd; }
    .trust-item strong { display: block; font-size: 13px; margin-bottom: 3px; }
    .steps { display: grid; gap: 10px; }
    .step { display: grid; grid-template-columns: 26px 1fr; gap: 10px; align-items: start; }
    .step-num { width: 26px; height: 26px; border-radius: 999px; background: var(--blue-soft); color: #174ea6; display: grid; place-items: center; font-size: 12px; font-weight: 850; }
    .results { display: none; overflow: hidden; }
    .results-head { display: flex; justify-content: space-between; gap: 16px; align-items: center; padding: 18px 20px; border-bottom: 1px solid var(--soft-line); }
    .summary-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; padding: 16px 20px 4px; }
    .metric { border: 1px solid var(--soft-line); border-radius: 8px; padding: 12px; background: #fbfcfd; min-height: 76px; }
    .metric .value { font-size: 22px; line-height: 1.1; font-weight: 850; }
    .metric .label { color: var(--muted); font-size: 12px; margin-top: 5px; line-height: 1.35; }
    .table-wrap { overflow-x: auto; padding: 12px 20px 20px; }
    table { width: 100%; border-collapse: collapse; font-size: 14px; min-width: 760px; }
    th, td { text-align: left; padding: 11px 10px; border-bottom: 1px solid var(--soft-line); vertical-align: top; }
    th { color: #485766; font-size: 12px; text-transform: uppercase; letter-spacing: .06em; background: #fbfcfd; }
    tbody tr:hover { background: #fbfcfd; }
    .course { font-weight: 800; }
    .course.cancelled { color: var(--error); text-decoration: line-through; }
    .meta { color: var(--muted); font-size: 13px; margin-top: 3px; }
    .location { display: inline-flex; align-items: center; min-height: 26px; border-radius: 999px; padding: 3px 9px; background: var(--green-soft); color: var(--green-dark); font-size: 12px; font-weight: 800; }
    .cancelled-pill { display: inline-flex; align-items: center; min-height: 24px; border-radius: 999px; padding: 2px 8px; background: var(--error-soft); color: var(--error); font-size: 12px; font-weight: 850; margin-left: 6px; }
    .assurance { margin-top: 14px; display: flex; flex-wrap: wrap; gap: 10px; color: var(--muted); font-size: 13px; }
    .assurance span { border: 1px solid var(--soft-line); border-radius: 999px; padding: 5px 9px; background: #fbfcfd; }
    .privacy-list { margin: 0; padding: 0; list-style: none; display: grid; gap: 11px; }
    .privacy-list li { display: grid; grid-template-columns: 8px 1fr; gap: 10px; color: var(--muted); font-size: 14px; line-height: 1.45; }
    .privacy-list li:before { content: ""; width: 8px; height: 8px; border-radius: 99px; background: var(--green); margin-top: 6px; }
    .notice { border: 1px solid #f2d38c; background: var(--warn-soft); color: var(--warn); border-radius: 8px; padding: 12px; font-size: 13px; line-height: 1.45; }
    .empty { padding: 22px 20px; color: var(--muted); border-top: 1px solid var(--soft-line); display: none; }
    .empty strong { color: var(--ink); display: block; margin-bottom: 4px; }
    .support { display: none; margin: 0 20px 20px; border-top: 1px solid var(--soft-line); padding-top: 16px; color: var(--muted); font-size: 13px; line-height: 1.45; }
    .support-main { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; }
    .support strong { color: var(--ink); }
    .support button, .support a.button { min-height: 34px; padding: 0 10px; font-size: 12px; }
    .support button.ghost { background: transparent; color: var(--green-dark); border: 1px solid var(--line); }
    .support button.ghost:hover { background: var(--green-soft); }
    .upi-row { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
    .upi-id { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace; color: var(--ink); background: #fff; border: 1px solid var(--line); border-radius: 7px; padding: 7px 9px; }
    .week-view { display: grid; gap: 12px; padding: 14px 20px 20px; }
    .routine-board { min-width: 960px; padding: 0; color: #1f2933; }
    .routine-title {
      display: flex;
      justify-content: space-between;
      gap: 20px;
      align-items: center;
      padding: 13px 14px;
      border: 1px solid var(--soft-line);
      border-radius: 8px;
      background: #fbfcfd;
      margin-bottom: 10px;
      font-size: 14px;
      font-weight: 850;
    }
    .routine-title span:last-child { color: var(--green-dark); font-size: 12px; }
    .routine-grid {
      display: grid;
      grid-template-columns: 72px repeat(7, minmax(112px, 1fr));
      gap: 8px;
      align-items: stretch;
    }
    .routine-head {
      min-height: 42px;
      display: grid;
      place-items: center;
      border: 1px solid var(--soft-line);
      border-radius: 8px;
      background: #f6f9fb;
      color: #344455;
      font-size: 13px;
      font-weight: 850;
    }
    .routine-head.today {
      background: var(--green-soft);
      color: var(--green-dark);
      border-color: #b9dfc9;
    }
    .routine-time {
      min-height: 86px;
      display: flex;
      align-items: flex-start;
      justify-content: flex-end;
      padding: 12px 8px 0 0;
      color: #465766;
      font-size: 12px;
      font-weight: 850;
      white-space: nowrap;
    }
    .routine-cell {
      min-height: 86px;
      border: 1px solid var(--soft-line);
      border-radius: 8px;
      background: #fff;
      padding: 8px;
      font-size: 13px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }
    .routine-cell.today { border-color: #c7e8d4; background: #fbfffd; }
    .routine-card {
      height: 100%;
      display: grid;
      align-content: start;
      gap: 4px;
      border-left: 4px solid var(--green);
      border-radius: 7px;
      background: #f8fbff;
      padding: 8px;
      box-shadow: inset 0 0 0 1px #e6eef6;
    }
    .routine-card strong { display: block; font-size: 13px; font-weight: 850; color: #14212b; }
    .routine-card span { display: block; color: #536273; font-size: 12px; }
    .routine-card .room { color: var(--green-dark); font-weight: 800; }
    .routine-empty {
      display: grid;
      place-items: center;
      color: #8793a0;
      background: #fafbfc;
      border-style: dashed;
      font-size: 12px;
      font-weight: 750;
    }
    .routine-mobile { display: none; }
    .mobile-day { border: 1px solid var(--soft-line); border-radius: 8px; background: #fff; padding: 11px; display: grid; gap: 9px; }
    .mobile-day.today { border-color: #b9dfc9; background: #fbfffd; }
    .mobile-day-head { display: flex; align-items: center; justify-content: space-between; gap: 10px; font-weight: 850; }
    .mobile-day-head span { color: var(--muted); font-size: 12px; font-weight: 750; }
    .mobile-slot { display: grid; grid-template-columns: 76px 1fr; gap: 10px; border: 1px solid var(--soft-line); border-left: 4px solid var(--green); border-radius: 8px; background: #f8fbff; padding: 10px; }
    .mobile-slot-time { color: #344455; font-weight: 850; font-size: 13px; line-height: 1.35; }
    .mobile-slot strong { display: block; font-size: 14px; line-height: 1.35; }
    .mobile-slot span { display: block; color: var(--muted); font-size: 12px; margin-top: 2px; }
    .day-band { border-top: 1px solid var(--soft-line); padding-top: 12px; display: grid; gap: 8px; }
    .day-title { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; color: var(--ink); font-weight: 850; }
    .day-title span { color: var(--muted); font-size: 12px; font-weight: 750; }
    .day-events { display: grid; gap: 8px; }
    .week-event { display: grid; grid-template-columns: 108px 1fr auto; gap: 10px; align-items: start; border: 1px solid var(--soft-line); border-radius: 8px; padding: 10px; background: #fbfcfd; }
    .week-empty { border: 1px dashed var(--line); border-radius: 8px; padding: 10px; color: var(--muted); background: #fff; font-size: 13px; }
    .week-time { color: #344455; font-size: 13px; font-weight: 850; white-space: nowrap; }
    .week-course { min-width: 0; }
    .week-course strong { display: block; font-size: 13px; line-height: 1.35; }
    .week-course .meta { margin-top: 2px; }
    .week-event.cancelled { background: #fff8f8; border-color: #f4caca; }

    @media (max-width: 520px) {
      body { background: #f8fafb; }
      .site-nav { padding: 14px 14px 12px; align-items: center; background: #fff; border-bottom: 1px solid var(--soft-line); }
      .brand-logo { width: 112px; }
      .nav-actions { flex: 0 0 auto; }
      .hero { min-height: 112px; padding: 22px 14px; background-position: center 35%; }
      .hero h1 { font-size: 34px; line-height: 1.02; }
      main { padding: 16px 10px 36px; }
      .layout, .sync-card { gap: 12px; }
      .panel { border-radius: 8px; box-shadow: none; }
      .panel-pad { padding: 16px; }
      .form-grid { grid-template-columns: 1fr; gap: 9px; }
      input { min-height: 46px; font-size: 16px; }
      button, a.button { width: 100%; }
      .nav-actions .button { width: auto; }
      .assurance { gap: 6px; font-size: 12px; }
      .assurance span { padding: 4px 7px; }
      .trust-row, .summary-grid { grid-template-columns: 1fr; }
      .title-row, .results-head { flex-direction: column; align-items: stretch; }
      .results-head { padding: 15px 16px; }
      .summary-grid { padding: 12px 16px 0; }
      .metric { min-height: 64px; padding: 10px; }
      .metric .value { font-size: 20px; }
      .table-wrap { padding: 8px 16px 16px; }
      .week-event { grid-template-columns: 1fr; }
      .routine-board { min-width: 0; }
      .routine-title { font-size: 15px; margin-bottom: 4px; padding-bottom: 10px; border-bottom-width: 1px; }
      .routine-grid { display: none; }
      .routine-mobile { display: grid; gap: 2px; }
      .week-view { padding: 10px 16px 16px; gap: 8px; }
      .support { margin: 0 16px 16px; }
      .upi-row { align-items: stretch; }
      .upi-id { width: 100%; text-align: center; }
    }
  </style>
</head>
<body>
  <header class="site-nav">
    <strong>Campus Calendar Sync</strong>
    <div class="nav-actions">
      {% if authed %}
        <a class="button" href="/logout">Sign out</a>
      {% else %}
        <a class="button" href="/authorize">Sign in</a>
      {% endif %}
    </div>
  </header>

  <section class="hero">
    <h1>The Campus Calendar Syncer</h1>
  </section>

  <main>
    <div class="layout">
      <div class="sync-card">
        <section class="panel panel-pad">
          <div class="title-row">
            <div>
              <div class="eyebrow">Live sheet connected</div>
              <h2>Roll number search</h2>
              <div class="muted">Check your classes, then sync them in one tap.</div>
            </div>
            <div class="status-chip"><span class="dot"></span>Green label enabled</div>
          </div>

          {% if not authed %}
            <div class="trust-row" style="margin-top:18px">
              <div class="trust-item"><strong>Student owned</strong><span class="muted">Events go to the Google account that signs in.</span></div>
              <div class="trust-item"><strong>No duplicates</strong><span class="muted">Re-sync updates existing class events.</span></div>
              <div class="trust-item"><strong>Live timetable</strong><span class="muted">Uses the current timetable file before syncing.</span></div>
            </div>
            <div style="margin-top:18px">
              <a class="button" href="/authorize">Sign in with Google</a>
            </div>
          {% else %}
            <div class="form-grid" style="margin-top:18px">
              <div>
                <label for="student">Your roll number or name</label>
                <input id="student" value="{{ guessed_identifier }}" placeholder="Example: M29B-27 or Mahesh Babu" autocomplete="off">
              </div>
              <button id="previewBtn" class="secondary" onclick="preview()">Find classes</button>
              <button id="syncBtn" onclick="sync()">Sync calendar</button>
            </div>
            <div style="margin-top:10px">
              <button id="clearBtn" class="danger" onclick="clearSync()">Clear sync</button>
            </div>
            <div class="assurance" id="assurance">
              <span>Reads live sheet every click</span>
              <span>Updates existing events</span>
              <span>Shows cancelled classes</span>
            </div>
            <div class="status" id="status"></div>
          {% endif %}
        </section>

        <section class="panel results" id="results">
          <div class="results-head">
            <div>
              <h2>Matched classes</h2>
              <div class="muted" id="studentSummary"></div>
            </div>
            <div class="status-chip" id="resultChip"><span class="dot"></span>Ready to sync</div>
          </div>
          <div class="summary-grid" id="metrics"></div>
          <div class="empty" id="emptyState">
            <strong>No future matching classes found.</strong>
            This usually means the student ID/name does not match the section workbook, or the live timetable has not published those classes yet.
          </div>
          <div class="table-wrap">
            <table id="eventsTable">
              <thead id="eventsHead"><tr><th>Date</th><th>Time</th><th>Class</th><th>Room</th></tr></thead>
              <tbody id="events"></tbody>
            </table>
            <div class="week-view" id="weekView"></div>
          </div>
          {% if support.enabled %}
            <div class="support" id="supportBox">
              <div class="support-main">
                <span><strong>No more "wait, do I have Strategy or FRM right now" moments.</strong> This runs on a free-tier server that I keep alive with mild optimism and zero budget. If it saved you a few minutes of squinting at the timetable sheet,</span>
                <a class="button" href="{{ support.upi_url }}">buy it a coffee</a>
                <span>no pressure, it'll survive either way.</span>
              </div>
            </div>
          {% endif %}
        </section>
      </div>

    </div>
  </main>

  <script>
    const previewBtn = () => document.getElementById('previewBtn');
    const syncBtn = () => document.getElementById('syncBtn');
    const clearBtn = () => document.getElementById('clearBtn');

    async function post(path) {
      if (path === '/sync') {
        await syncInChunks();
        return;
      }
      const identifier = document.getElementById('student').value.trim();
      const status = document.getElementById('status');
      status.className = 'status';
      status.textContent = identifier ? 'Checking the latest timetable...' : 'Checking the latest timetable for your signed-in email...';
      setBusy(true);
      status.classList.add('loading');
      let data = {};
      try {
        const res = await fetch(path, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ identifier })
        });
        data = await res.json();
        if (!res.ok) {
          status.textContent = friendlyError(data.error);
          status.className = 'status error';
          return;
        }
      } catch (error) {
        status.textContent = 'Connection interrupted. Please try again in a moment.';
        status.className = 'status error';
        return;
      } finally {
        setBusy(false);
      }
      render(data, path);
      status.textContent = data.message || 'Preview ready.';
      status.className = 'status ok';
    }

    async function clearSyncedEvents() {
      const identifier = document.getElementById('student').value.trim();
      const status = document.getElementById('status');
      if (!window.confirm('Clear class events created by this app for this student?')) return;
      status.className = 'status';
      status.textContent = 'Clearing synced classes...';
      setBusy(true);
      status.classList.add('loading');
      let data = {};
      try {
        const res = await fetch('/clear-sync', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ identifier })
        });
        data = await res.json();
        if (!res.ok) {
          status.textContent = friendlyError(data.error);
          status.className = 'status error';
          return;
        }
      } catch (error) {
        status.textContent = 'Connection interrupted. Please try again in a moment.';
        status.className = 'status error';
        return;
      } finally {
        setBusy(false);
      }
      render(data, '/clear-sync');
      status.textContent = data.message || 'Synced classes cleared.';
      status.className = 'status ok';
    }

    async function syncInChunks() {
      const identifier = document.getElementById('student').value.trim();
      const status = document.getElementById('status');
      status.className = 'status';
      status.textContent = identifier ? 'Starting calendar sync...' : 'Starting calendar sync for your signed-in email...';

      setBusy(true);
      status.classList.add('loading');
      let offset = 0;
      let finalData = null;
      const chunkSize = 8;
      try {
        while (true) {
          const res = await fetch('/sync', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ identifier, offset, limit: chunkSize })
          });
          const data = await res.json();
          if (!res.ok) {
            status.textContent = friendlyError(data.error);
            status.className = 'status error';
            return;
          }
          finalData = data;
          const total = data.total_events || data.events.length || 0;
          const synced = Math.min(data.next_offset || total, total);
          status.textContent = total ? `Syncing calendar... ${synced}/${total}` : 'No matching classes to sync.';
          if (data.done) break;
          offset = data.next_offset;
        }
      } catch (error) {
        status.textContent = 'Connection interrupted. Please keep this page open and try again.';
        status.className = 'status error';
        return;
      } finally {
        setBusy(false);
      }
      render(finalData, '/sync');
      status.textContent = 'Calendar synced.';
      status.className = 'status ok';
    }

    function render(data, path) {
      document.getElementById('results').style.display = 'block';
      const support = document.getElementById('supportBox');
      const assurance = document.getElementById('assurance');
      const resultChip = document.getElementById('resultChip');
      if (support) support.style.display = path === '/sync' ? 'block' : 'none';
      if (assurance) assurance.style.display = path === '/sync' ? 'none' : 'flex';
      if (resultChip) resultChip.style.display = path === '/sync' ? 'none' : 'inline-flex';
      document.querySelector('#results h2').textContent =
        path === '/preview' ? 'Matched classes' : path === '/clear-sync' ? 'Sync cleared' : 'Calendar synced';
      document.getElementById('studentSummary').textContent =
        path === '/preview'
          ? `${data.student.name} (${data.student.student_id}) - ${courseSummary(data.events).length} matched courses`
          : `${data.student.name} (${data.student.student_id})`;
      const metrics = document.getElementById('metrics');
      const tbody = document.getElementById('events');
      const table = document.getElementById('eventsTable');
      const weekView = document.getElementById('weekView');
      const empty = document.getElementById('emptyState');
      tbody.innerHTML = '';
      if (weekView) weekView.innerHTML = '';
      empty.style.display = data.events.length ? 'none' : 'block';
      if (path === '/preview') {
        if (metrics) metrics.style.display = 'grid';
        renderMetrics(data.events);
        if (table) table.style.display = 'none';
        renderWeekView(data.events, weekView);
        return;
      }
      if (metrics) metrics.style.display = 'none';
      empty.style.display = 'none';
      if (table) table.style.display = 'none';
      document.getElementById('eventsHead').innerHTML = '<tr><th>Date</th><th>Time</th><th>Class</th><th>Room</th></tr>';
    }

    function renderWeekView(events, container) {
      if (!container) return;
      const grouped = regularWeeklyPattern(events);
      const pattern = Array.from(grouped.values()).flat();
      const timeSlots = Array.from(new Set(pattern.map(event => event.start_time))).sort();
      const section = pattern[0]?.section || '';
      const board = document.createElement('div');
      board.className = 'routine-board';
      board.innerHTML = `
        <div class="routine-title"><span>${section ? `Section ${section} · ` : ''}Term IV</span><span>[Today: ${todayShortDayName()}]</span></div>
        <div class="routine-grid"></div>
        <div class="routine-mobile"></div>
      `;
      const grid = board.querySelector('.routine-grid');
      const mobile = board.querySelector('.routine-mobile');
      const todayIndex = todayDayIndex();
      grid.appendChild(document.createElement('div'));
      for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
        const head = document.createElement('div');
        head.className = `routine-head${dayIndex === todayIndex ? ' today' : ''}`;
        head.textContent = shortDayName(dayIndex);
        grid.appendChild(head);
      }
      for (const time of timeSlots) {
        const timeCell = document.createElement('div');
        timeCell.className = 'routine-time';
        timeCell.textContent = time;
        grid.appendChild(timeCell);
        for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
          const event = (grouped.get(dayIndex) || []).find(item => item.start_time === time);
          const cell = document.createElement('div');
          cell.className = `routine-cell${dayIndex === todayIndex ? ' today' : ''}`;
          if (event) {
            cell.innerHTML = `
              <div class="routine-card">
                <strong>${shortSubject(event.subject)}</strong>
                ${professorName(event.professor) ? `<span>${professorName(event.professor)}</span>` : ''}
                ${event.location ? `<span class="room">${event.location}</span>` : ''}
              </div>
            `;
          }
          grid.appendChild(cell);
        }
      }
      if (!timeSlots.length) {
        const empty = document.createElement('div');
        empty.className = 'week-empty';
        empty.textContent = 'No regular weekly pattern found yet.';
        board.appendChild(empty);
      } else {
        addNoClassMarkers(grid, grouped, timeSlots.length);
        renderMobileRoutine(mobile, grouped);
      }
      container.appendChild(board);
    }

    function courseSummary(events) {
      const grouped = new Map();
      for (const event of events) {
        const key = `${event.course_code}|${event.section}`;
        if (!grouped.has(key)) {
          grouped.set(key, {
            course_code: event.course_code,
            subject: event.subject,
            section: event.section,
            total: 0,
            cancelled: 0
          });
        }
        const item = grouped.get(key);
        item.total += 1;
        if (event.cancelled) item.cancelled += 1;
      }
      return Array.from(grouped.values()).sort((a, b) => a.course_code.localeCompare(b.course_code));
    }

    function regularWeeklyPattern(events) {
      const grouped = new Map();
      for (const event of events) {
        if (!isRoutineEvent(event)) continue;
        const dayIndex = weekdayIndex(event.date);
        const key = [dayIndex, event.start_time, event.end_time, event.course_code, event.section].join('|');
        if (!grouped.has(key)) {
          grouped.set(key, { ...event, dayIndex, count: 0 });
        }
        grouped.get(key).count += 1;
      }
      const repeating = Array.from(grouped.values()).filter(event => event.count >= 2);
      const byDay = new Map();
      for (const event of repeating) {
        if (!byDay.has(event.dayIndex)) byDay.set(event.dayIndex, []);
        byDay.get(event.dayIndex).push(event);
      }
      for (const dayEvents of byDay.values()) {
        dayEvents.sort((a, b) => a.start_time.localeCompare(b.start_time));
      }
      return byDay;
    }

    function isRoutineEvent(event) {
      if (event.cancelled) return false;
      const text = `${event.subject || ''} ${event.location || ''} ${event.professor || ''} ${event.source_sheet || ''}`.toLowerCase();
      return !text.includes('reschedule') && !text.includes('rescheduled');
    }

    function weekdayIndex(value) {
      const day = new Date(`${value}T00:00:00`).getDay();
      return (day + 6) % 7;
    }

    function dayName(index) {
      return ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'][index];
    }

    function shortDayName(index) {
      return ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'][index];
    }

    function todayShortDayName() {
      return new Intl.DateTimeFormat('en-IN', { weekday: 'short', timeZone: 'Asia/Kolkata' }).format(new Date());
    }

    function todayDayIndex() {
      const day = new Intl.DateTimeFormat('en-US', { weekday: 'short', timeZone: 'Asia/Kolkata' }).format(new Date());
      return ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].indexOf(day);
    }

    function addNoClassMarkers(grid, grouped, rowCount) {
      for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
        if ((grouped.get(dayIndex) || []).length) continue;
        const cellIndex = 9 + dayIndex;
        const cell = grid.children[cellIndex];
        if (cell && rowCount) {
          cell.className = 'routine-cell routine-empty';
          cell.textContent = 'No class';
        }
      }
    }

    function renderMobileRoutine(container, grouped) {
      if (!container) return;
      for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
        const dayEvents = grouped.get(dayIndex) || [];
        const day = document.createElement('section');
        day.className = `mobile-day${dayIndex === todayDayIndex() ? ' today' : ''}`;
        day.innerHTML = `<div class="mobile-day-head">${dayName(dayIndex)}<span>${dayEvents.length ? `${dayEvents.length} regular ${dayEvents.length === 1 ? 'slot' : 'slots'}` : 'No class'}</span></div>`;
        if (!dayEvents.length) {
          const empty = document.createElement('div');
          empty.className = 'week-empty';
          empty.textContent = 'No class';
          day.appendChild(empty);
          container.appendChild(day);
          continue;
        }
        for (const event of dayEvents) {
          const slot = document.createElement('article');
          slot.className = 'mobile-slot';
          slot.innerHTML = `
            <div class="mobile-slot-time">${event.start_time}</div>
            <div>
              <strong>${shortSubject(event.subject)}</strong>
              <span>${professorName(event.professor)}${event.location ? ` · ${event.location}` : ''}</span>
            </div>
          `;
          day.appendChild(slot);
        }
        container.appendChild(day);
      }
    }

    function shortSubject(value) {
      const clean = String(value || '').replace(/:.*$/, '').trim();
      const words = clean.split(/\\s+/).filter(Boolean);
      if (words.length <= 3) return clean;
      return words.slice(0, 3).join(' ');
    }

    function professorName(value) {
      const clean = String(value || '').replace(/^prof\\.?\\s*/i, '').trim();
      if (!clean) return '';
      const parts = clean.split(/\\s+/);
      return parts[parts.length - 1];
    }

    function renderMetrics(events) {
      const metrics = document.getElementById('metrics');
      const subjects = new Set(events.map(event => event.course_code)).size;
      const next = events[0] ? `${formatDate(events[0].date)} · ${events[0].start_time}` : 'No class';
      metrics.innerHTML = `
        <div class="metric"><div class="value">${events.length}</div><div class="label">upcoming calendar events</div></div>
        <div class="metric"><div class="value">${subjects}</div><div class="label">matched subjects</div></div>
        <div class="metric"><div class="value" style="font-size:16px">${next}</div><div class="label">next matched class</div></div>
      `;
    }

    function setBusy(isBusy) {
      if (previewBtn()) previewBtn().disabled = isBusy;
      if (syncBtn()) syncBtn().disabled = isBusy;
      if (clearBtn()) clearBtn().disabled = isBusy;
    }

    function friendlyError(raw) {
      const message = raw || '';
      if (message.includes('No matching student')) return 'No student found. Try student ID exactly as shown in the section list, for example M29B-27.';
      if (message.includes('Could not match your signed-in email')) return 'Could not auto-detect you from email. Enter your roll number once, for example M29B-27.';
      if (message.includes('Multiple students')) return `${message} Use the student ID to pick the right person.`;
      if (message.includes('Sign in')) return 'Sign in with Google first, then preview your classes.';
      if (message.includes('Sync already in progress')) return 'Sync already in progress for this student. Please wait for it to finish.';
      if (message.includes('Daily sync limit reached')) return message;
      return 'Could not read the timetable right now. Try again in a minute.';
    }

    function formatDate(value) {
      const date = new Date(`${value}T00:00:00`);
      return date.toLocaleDateString('en-IN', { weekday: 'short', day: '2-digit', month: 'short' });
    }

    function preview() { post('/preview'); }
    function sync() { post('/sync'); }
    function clearSync() { clearSyncedEvents(); }

    {% if support_preview %}
      window.addEventListener('DOMContentLoaded', () => {
        render({{ support_preview | safe }}, '/sync');
        const status = document.getElementById('status');
        if (status) {
          status.textContent = {{ support_preview_message | tojson }};
          status.className = 'status ok';
        }
      });
    {% endif %}
  </script>
</body>
</html>
"""


def credentials_from_session():
    encrypted = session.get("credentials")
    if not encrypted:
        return None
    if isinstance(encrypted, dict):
        return Credentials(**encrypted)
    try:
        token = json.loads(session_cipher().decrypt(encrypted.encode("utf-8")).decode("utf-8"))
    except (InvalidToken, ValueError, TypeError, json.JSONDecodeError):
        session.clear()
        return None
    return Credentials(**token)


def credentials_to_session(credentials):
    token = {
        "token": credentials.token,
        "refresh_token": credentials.refresh_token,
        "token_uri": credentials.token_uri,
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
        "scopes": credentials.scopes,
    }
    session["credentials"] = session_cipher().encrypt(json.dumps(token).encode("utf-8")).decode("utf-8")


def session_cipher():
    key = hashlib.sha256(app.secret_key.encode("utf-8")).digest()
    return Fernet(base64.urlsafe_b64encode(key))


def require_config():
    if not get_client_secret_file():
        raise RuntimeError("Set GOOGLE_CLIENT_SECRET_FILE or GOOGLE_CLIENT_SECRET_JSON.")
    if not os.path.exists(WORKBOOK_PATH):
        raise RuntimeError("Set SECTION_WORKBOOK to the section distribution workbook path.")


def get_client_secret_file():
    global _client_secret_tempfile
    if CLIENT_SECRET_FILE and os.path.exists(CLIENT_SECRET_FILE):
        return CLIENT_SECRET_FILE
    if CLIENT_SECRET_JSON:
        if not _client_secret_tempfile:
            temp = NamedTemporaryFile("w", suffix=".json", delete=False)
            temp.write(CLIENT_SECRET_JSON)
            temp.close()
            _client_secret_tempfile = temp.name
        return _client_secret_tempfile
    return ""


def oauth_redirect_uri():
    return OAUTH_REDIRECT_URI or url_for("oauth2callback", _external=True)


def email_to_identifier(email: str) -> str:
    return ""


def guessed_identifier() -> str:
    return email_to_identifier(session.get("user_email", ""))


def resolve_student(sections: SectionRepository, identifier: str) -> tuple:
    entered = (identifier or "").strip()
    if not entered:
        raise ValueError("Enter your roll number or name first.")
    return sections.find_student(entered), entered


def email_from_id_token(id_token: str | None):
    if not id_token:
        return None
    try:
        payload = id_token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload.encode("utf-8")).decode("utf-8"))
    except (IndexError, ValueError, TypeError, json.JSONDecodeError):
        return None
    return (claims.get("email") or "").strip().lower() or None


def store_user_email(credentials):
    email = email_from_id_token(getattr(credentials, "id_token", None))
    if email:
        session["user_email"] = email
        return email

    req = url_request.Request(
        "https://www.googleapis.com/oauth2/v2/userinfo",
        headers={"Authorization": f"Bearer {credentials.token}"},
    )
    try:
        with url_request.urlopen(req, timeout=10) as response:
            profile = json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        app.logger.warning("Could not read signed-in Google email: %s", exc)
        return None
    email = (profile.get("email") or "").strip().lower()
    if email:
        session["user_email"] = email
    return email


def support_config():
    if not SUPPORT_UPI_ID:
        return {"enabled": False}
    query = urlencode(
        {
            "pa": SUPPORT_UPI_ID,
            "pn": SUPPORT_PAYEE_NAME,
            "tn": SUPPORT_NOTE,
            "cu": "INR",
        }
    )
    return {
        "enabled": True,
        "upi_id": SUPPORT_UPI_ID,
        "upi_url": f"upi://pay?{query}",
    }


def build_services():
    credentials = credentials_from_session()
    if not credentials:
        raise RuntimeError("Sign in with Google first.")
    sections = SectionRepository(WORKBOOK_PATH)
    timetable = TimetableService(SHEET_ID, credentials)
    calendar = CalendarSyncService(credentials)
    return sections, timetable, calendar


def sync_lock_key(student):
    credentials = credentials_from_session()
    token_fingerprint = hashlib.sha256(
        (getattr(credentials, "refresh_token", "") or getattr(credentials, "token", "") or "anonymous").encode("utf-8")
    ).hexdigest()[:16]
    return f"{token_fingerprint}:{student.student_id.upper()}"


def acquire_sync_lock(key: str):
    with _sync_locks_guard:
        lock = _sync_locks.setdefault(key, Lock())
    return lock, lock.acquire(blocking=False)


def sync_user_key(student):
    email = (session.get("user_email") or "").strip().lower()
    if email:
        return f"email:{email}"
    raise RuntimeError("Could not verify your signed-in Google email. Please sign out and sign in again.")


def anonymous_user_hash():
    email = (session.get("user_email") or "").strip().lower()
    if not email:
        return "anonymous"
    salt = app.secret_key.encode("utf-8")
    return hmac.new(salt, email.encode("utf-8"), hashlib.sha256).hexdigest()[:16]


def log_usage(action: str, status: str, **details):
    safe_details = " ".join(f"{key}={value}" for key, value in sorted(details.items()) if value is not None)
    app.logger.info(
        "usage action=%s status=%s user_hash=%s %s",
        action,
        status,
        anonymous_user_hash(),
        safe_details,
    )


def consume_daily_sync_attempt(student):
    allowed, attempts = check_and_increment_sync(sync_user_key(student), MAX_DAILY_SYNC_ATTEMPTS)
    if allowed:
        return None
    log_usage("quota", "blocked", attempts=attempts, limit=MAX_DAILY_SYNC_ATTEMPTS)
    return (
        jsonify(
            {
                "error": (
                    f"Daily sync limit reached. You can sync or clear up to "
                    f"{MAX_DAILY_SYNC_ATTEMPTS} times per day."
                )
            }
        ),
        429,
    )


@app.route("/")
def index():
    return render_template_string(
        PAGE,
        authed=bool(credentials_from_session()),
        guessed_identifier=guessed_identifier(),
        support=support_config(),
        support_preview="",
        support_preview_message="",
    )


@app.route("/support-preview")
def support_preview():
    if not ENABLE_SUPPORT_PREVIEW:
        return redirect(url_for("index"))
    demo = {
        "student": {"name": "Preview Student", "student_id": "M31C-25"},
        "events": [
            {
                "date": "2026-07-02",
                "start_time": "09:00",
                "end_time": "10:15",
                "course_code": "26MBA101",
                "subject": "Strategy",
                "section": "A",
                "location": "CR-1",
                "professor": "Prof. Preview",
                "source_sheet": "Preview Week",
                "cancelled": False,
            },
            {
                "date": "2026-07-03",
                "start_time": "11:30",
                "end_time": "12:45",
                "course_code": "26MBA202",
                "subject": "Financial Risk Management",
                "section": "A",
                "location": "CR-2",
                "professor": "",
                "source_sheet": "Preview Week",
                "cancelled": True,
            },
        ],
    }
    return render_template_string(
        PAGE,
        authed=True,
        guessed_identifier="Preview Student",
        support=support_config(),
        support_preview=json.dumps(demo),
        support_preview_message="Preview: 1 new and 1 updated event.",
    )


@app.route("/health")
@app.route("/healthz")
@app.route("/healthz/")
def healthz():
    return jsonify({"status": "ok"})


@app.route("/authorize")
def authorize():
    require_config()
    flow = Flow.from_client_secrets_file(get_client_secret_file(), scopes=SCOPES)
    flow.redirect_uri = oauth_redirect_uri()
    auth_url, state = flow.authorization_url(
        access_type="offline",
        prompt="consent",
    )
    session["state"] = state
    return redirect(auth_url)


@app.route("/oauth2callback")
def oauth2callback():
    require_config()
    if "state" not in session:
        return redirect(url_for("logout"))

    flow = Flow.from_client_secrets_file(get_client_secret_file(), scopes=SCOPES, state=session["state"])
    flow.redirect_uri = oauth_redirect_uri()
    try:
        flow.fetch_token(authorization_response=request.url)
    except MismatchingStateError:
        app.logger.warning("OAuth state mismatch; clearing session so the user can sign in again.")
        session.clear()
        return redirect(url_for("index"))
    credentials_to_session(flow.credentials)
    store_user_email(flow.credentials)
    return redirect(url_for("index"))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.route("/preview", methods=["POST"])
def preview():
    identifier = ""
    try:
        identifier = request.json.get("identifier", "")
        sections, timetable, _ = build_services()
        student, identifier = resolve_student(sections, identifier)
        events = timetable.match_events(student.courses)
        log_usage("find", "success", matched_events=len(events), matched_courses=len(student.courses))
        return jsonify(
            {
                "student": asdict(student),
                "events": [asdict(event) for event in events],
                "message": f"Found {len(events)} matching classes.",
            }
        )
    except Exception as exc:
        log_usage("find", "failed")
        app.logger.exception("preview_failed error=%s", exc)
        return jsonify({"error": str(exc)}), 400


@app.route("/sync", methods=["POST"])
def sync():
    identifier = ""
    try:
        identifier = request.json.get("identifier", "")
        offset = max(int(request.json.get("offset", 0) or 0), 0)
        limit = request.json.get("limit")
        limit = max(min(int(limit), 25), 1) if limit is not None else None
        sections, timetable, calendar = build_services()
        student, identifier = resolve_student(sections, identifier)
        lock, acquired = acquire_sync_lock(sync_lock_key(student))
        if not acquired:
            log_usage("sync", "blocked", reason="in_progress")
            return jsonify({"error": "Sync already in progress for this student."}), 409
        try:
            if offset == 0:
                quota_response = consume_daily_sync_attempt(student)
                if quota_response:
                    return quota_response
                log_usage("sync", "accepted")
            events = timetable.match_events(student.courses)
            chunk = events[offset : offset + limit] if limit else events
            result = calendar.upsert_events(student, chunk)
        finally:
            lock.release()
        next_offset = offset + len(chunk)
        done = next_offset >= len(events)
        if done:
            log_usage("sync", "success", matched_events=len(events))
        return jsonify(
            {
                "student": asdict(student),
                "events": [asdict(event) for event in events],
                "message": f"Synced {result['created']} new and {result['updated']} updated events.",
                "total_events": len(events),
                "next_offset": next_offset,
                "done": done,
            }
        )
    except Exception as exc:
        log_usage("sync", "failed")
        app.logger.exception("sync_failed error=%s", exc)
        return jsonify({"error": str(exc)}), 400


@app.route("/clear-sync", methods=["POST"])
def clear_sync():
    identifier = ""
    try:
        identifier = request.json.get("identifier", "")
        sections, timetable, calendar = build_services()
        student, identifier = resolve_student(sections, identifier)
        lock, acquired = acquire_sync_lock(sync_lock_key(student))
        if not acquired:
            log_usage("clear", "blocked", reason="in_progress")
            return jsonify({"error": "Sync already in progress for this student."}), 409
        try:
            quota_response = consume_daily_sync_attempt(student)
            if quota_response:
                return quota_response
            events = timetable.match_events(student.courses)
            result = calendar.delete_events(student, events)
        finally:
            lock.release()
        log_usage("clear", "success", matched_events=len(events), deleted=result["deleted"], missing=result["missing"])
        return jsonify(
            {
                "student": asdict(student),
                "events": [asdict(event) for event in events],
                "message": f"Cleared {result['deleted']} synced classes from Google Calendar.",
                "deleted": result["deleted"],
                "missing": result["missing"],
            }
        )
    except Exception as exc:
        log_usage("clear", "failed")
        app.logger.exception("clear_sync_failed error=%s", exc)
        return jsonify({"error": str(exc)}), 400


if __name__ == "__main__":
    os.environ.setdefault("OAUTHLIB_INSECURE_TRANSPORT", "1")
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5050")), debug=True)
