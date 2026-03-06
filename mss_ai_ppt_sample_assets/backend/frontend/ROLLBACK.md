# UI Rollback Guide

This document describes how to rollback the UI files in `backend/frontend` to a previous backup snapshot.

## Backup Location

- Directory pattern: `backend/frontend/__backup__/YYYYMMDD_HHMMSS`
- This release backup: `backend/frontend/__backup__/20260305_092944`

## Rollback Steps

1. Stop backend service.
2. Copy backup files back into `backend/frontend`.
3. Start backend service.
4. Verify:
   - `http://<host>:8000/ui/index.html`
   - `http://<host>:8000/ui/login.html`
   - `http://<host>:8000/ui/admin.html`

## Windows PowerShell Example

```powershell
$frontend = "mss_ai_ppt_sample_assets/backend/frontend"
$backup = "$frontend/__backup__/20260305_092944"

Copy-Item "$backup/index.html" "$frontend/index.html" -Force
Copy-Item "$backup/login.html" "$frontend/login.html" -Force
Copy-Item "$backup/admin.html" "$frontend/admin.html" -Force

Remove-Item "$frontend/i18n" -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item "$backup/i18n" "$frontend/i18n" -Recurse -Force

Remove-Item "$frontend/assets" -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item "$backup/assets" "$frontend/assets" -Recurse -Force
```

## Linux Shell Example

```bash
FRONTEND="mss_ai_ppt_sample_assets/backend/frontend"
BACKUP="$FRONTEND/__backup__/20260305_092944"

cp -f "$BACKUP/index.html" "$FRONTEND/index.html"
cp -f "$BACKUP/login.html" "$FRONTEND/login.html"
cp -f "$BACKUP/admin.html" "$FRONTEND/admin.html"

rm -rf "$FRONTEND/i18n" "$FRONTEND/assets"
cp -r "$BACKUP/i18n" "$FRONTEND/i18n"
cp -r "$BACKUP/assets" "$FRONTEND/assets"
```
