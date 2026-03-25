# Docker Image Build And Push

This repository already contains the compiled frontend assets under:

`mss_ai_ppt_sample_assets/backend/frontend/`

The Docker image uses those static files directly. If you changed anything in `ui-src/`, rebuild the frontend before committing:

```powershell
cd f:\report-generation\ui-src
npm install
npm run build
```

Then commit both the source changes and the updated files under `mss_ai_ppt_sample_assets/backend/frontend/`.

## Build On Internal Server

Run these commands on the internal server after pulling the repository:

```powershell
cd <project-dir>
docker build -t report-generation:20260325 .
```

## Push To Artifact Registry

```powershell
docker login docker.sangfor.com -u product_2573_676826 -p 2a17071e9c637d51
docker tag report-generation:20260325 docker.sangfor.com/mss-release/report-generation:20260325
docker push docker.sangfor.com/mss-release/report-generation:20260325
```

## Runtime Configuration

Do not bake secrets into the image. Inject these at runtime on the target environment:

- `ENABLE_LLM`
- `OPENAI_API_KEY`
- `OPENAI_BASE_URL`
- `OPENAI_MODEL`
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD_HASH`
- `ADMIN_SESSION_SECRET`
- `LIBREOFFICE_PATH` if `soffice` is not on `PATH`
