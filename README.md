# Timesheets

Timesheet management, approvals, expenses and job costing. Two ways to run it:

| | Hosted (online) | Desktop installer |
|---|---|---|
| For | Selling to companies — they sign up on a web address, nothing to install | A single company that wants it on their own PCs |
| Entry point | `web/server.py` | `timesheet_app.py` via `Build Setup.bat` |
| Data | One private database per company on the server, daily backups | `%APPDATA%\Timesheets` or a shared folder |

Both use the same `timesheet_app.py` — all screens and rules live there once.

## Hosted version

**How a customer uses it:** open the site → "Create your company" → they're the
admin → add their team under Users (name, email, starting password) → staff sign
in from any browser with email + password. Admins change company name, logo and
targets under Settings. Each company is fully isolated.

**Run it on your PC to try it:** double-click `Run Web Locally.bat`, open
http://localhost:8000.

**Put it live (Render, about 10 minutes):**
1. Push this folder to a GitHub repo.
2. On render.com: New → Blueprint → pick the repo. `render.yaml` sets everything
   up (Docker build, 5 GB persistent disk at `/data`, random `SECRET_KEY`).
3. Add your domain (e.g. `timesheets.kivo.co.uk`) under the service's Settings →
   Custom Domains. Render gives you HTTPS automatically.

`fly.toml` is there if you'd rather use Fly.io; any host that runs a Docker
container with a persistent volume works (`DATA_ROOT=/data`, `SECURE_COOKIES=1`,
`SECRET_KEY=<long random string>`).

**What's in the box server-side:** `web/server.py` wraps the app's API over
HTTP; signed-cookie sessions (1 day, or 30 with "keep me signed in"); browser
uploads for receipts/bills/imports; exports come back as downloads; an accounts
index maps each email to its company; backups every 6 h (newest 30 kept per
company); exports are deleted after 6 h.

**Not in v1 (worth adding before scale):** self-service password reset (an
admin resets passwords from Users for now), email invites, billing/Stripe,
rate limiting. Postgres isn't needed until you have hundreds of companies.

## Desktop installer

See `Build Setup.bat` — builds `Output\Timesheets-Setup-1.0.0.exe` with
PyInstaller + Inno Setup. First run shows a setup screen; Settings has
"Move to a shared folder…" so a team can share one database on a network drive.

## Versioning

Bump `APP_VERSION` in `timesheet_app.py` (and `AppVersion` in `installer.iss`
for the desktop build).
