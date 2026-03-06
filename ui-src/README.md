# UI Source (React + Vite)

This directory contains the maintainable source code for `/ui/index.html`, `/ui/login.html`, and `/ui/admin.html`.

## Commands

```bash
npm install
npm run build
```

`npm run build` outputs static files directly to:

`mss_ai_ppt_sample_assets/backend/frontend/`

## Deployment model

- Build on Windows/dev/CI with Node.js.
- Deploy generated static files with backend code.
- CentOS runtime does not need Node.js/npm.
