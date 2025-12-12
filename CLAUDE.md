# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MSS AI PPT is a report generation platform that creates PowerPoint presentations from structured input data. The system supports multiple templates (management/technical), uses token-based placeholder replacement, and provides a web UI for slide editing and preview.

**Architecture**: Monorepo with React frontend + FastAPI backend

## Development Commands

### Backend (FastAPI)

```bash
# Run backend server from project root
cd mss_ai_ppt_sample_assets/backend
python -m uvicorn mss_ai_ppt_sample_assets.backend.app:app --reload --port 8000

# Alternative: run directly
cd mss_ai_ppt_sample_assets/backend
python app.py
```

### Frontend (React + Vite)

```bash
# Install dependencies
cd frontend
npm install

# Development server
npm run dev

# Build for production
npm run build

# Lint
npm run lint
```

## Code Architecture

### Backend Pipeline (mss_ai_ppt_sample_assets/backend/)

The backend follows a modular pipeline architecture:

1. **ReportService** ([services/report_service.py](mss_ai_ppt_sample_assets/backend/services/report_service.py)) - Orchestrates the entire report generation flow
   - Entry point for all operations (generate, rewrite, preview)
   - Coordinates between modules: data prep → LLM → validation → PPT generation → preview

2. **Core Pipeline Modules** (modules/):
   - **data_prep.py** - Transforms raw TenantInput into template-specific facts
   - **llm_orchestrator.py** - Generates SlideSpec (uses mock by default; real LLM optional)
   - **validator.py** - Schema validation and fact-checking against prepared data
   - **ppt_generator.py** - Token replacement in PPTX templates using python-pptx
   - **preview_generator.py** - Converts PPTX → PDF → PNG images for web preview
   - **template_loader.py** - Manages template catalog (PPTX + descriptor JSON pairs)
   - **audit_logger.py** - Appends structured logs to outputs/logs/audit.jsonl

3. **Data Models** (models/):
   - **inputs.py** - TenantInput schema (alerts, incidents, vulnerabilities, etc.)
   - **slidespec.py** - SlideSpec schema (intermediate representation before rendering)
   - **templates.py** - TemplateDescriptor schema (defines slides and render_map)
   - **audit.py** - AuditEvent schema for logging

4. **Configuration** ([config.py](mss_ai_ppt_sample_assets/backend/config.py)):
   - Directory paths (DATA_DIR, OUTPUTS_DIR, TEMPLATES_DIR, etc.)
   - Settings class for environment variables (OPENAI_API_KEY, ENABLE_LLM, etc.)

### Frontend Architecture (frontend/src/)

Single-page React app with Zustand state management:

- **App.tsx** - Main container: fetches templates/inputs, orchestrates generate/rewrite flows
- **store/useSlides.ts** - Zustand store holding slidespec, jobId, previews, loading/error state
- **api/client.ts** - Typed API client wrapper around fetch (listTemplates, generate, rewrite, preview, getLogs)
- **Components**:
  - **Toolbar** - Input/template dropdowns + Generate button
  - **SlideList** - Left sidebar showing slide thumbnails
  - **SlidePreview** - Center panel showing PNG preview
  - **SlideEditor** - Right sidebar JSON editor for rewriting slide content

### Key Data Flow

```
User selects input + template
  → POST /generate → ReportService.generate()
    → DataPrep.prepare_facts_for_template()
    → LLMOrchestrator.generate_slidespec() [uses mock by default]
    → Validator.validate_schema() + fact_check()
    → PPTGenerator.render() → writes PPTX to outputs/reports/
    → Saves SlideSpec to outputs/slidespecs/
  → POST /preview → PPTPreviewGenerator.to_images()
    → Converts PPTX → PDF → PNG images via LibreOffice/Poppler
  → Frontend displays slide previews + JSON editor
  → User edits JSON → POST /rewrite → LLMOrchestrator.rewrite_slide()
    → Merges new_content into existing slide data
    → Re-renders PPTX and regenerates previews
```

### Template System

Each template consists of two files in `data/templates/`:
- **PPTX file** - PowerPoint with `{TOKEN}` placeholders in text
- **descriptor JSON** - Defines slides array with slide_no, slide_key, render_map
  - render_map: maps TOKEN → dotted path into SlideSpec data (e.g., "TITLE" → "summary.title")

Current templates:
- `mss_management_light_v1` - 5 slides (cover, executive_summary, alerts_overview, major_incidents, recommendations)
- `mss_technical_dark_v1` - 6 slides (cover, alerts_detail, incident_timeline, vulnerability_exposure, cloud_attack_surface, appendix_evidence)

### Environment Variables

```bash
# Backend (.env or shell)
OPENAI_API_KEY=sk-...          # Required if ENABLE_LLM=true
OPENAI_BASE_URL=https://...    # Optional: custom endpoint
OPENAI_MODEL=gpt-4o-mini       # Default model
ENABLE_LLM=true                # Enable real LLM (default: false, uses mock)
DEFAULT_LOCALE=zh-CN           # Default locale for content
```

### Important File Locations

- **Input data catalog**: `data/inputs/catalog.json` - defines available input datasets
- **Mock outputs**: `data/mock_outputs/` - pre-generated SlideSpecs for testing without LLM
- **Generated reports**: `outputs/reports/{input_id}_{template_id}.pptx`
- **Saved SlideSpecs**: `outputs/slidespecs/{input_id}_{template_id}.json`
- **Preview images**: `outputs/previews/{sanitized_job_id}/slide-*.png`
- **Audit logs**: `outputs/logs/audit.jsonl`

### Preview Generation Fallback Strategy

Preview generation attempts multiple methods (see [preview_generator.py](mss_ai_ppt_sample_assets/backend/modules/preview_generator.py)):
1. PPTX → PDF (LibreOffice headless)
2. PDF → PNG (Poppler pdf2image or PyMuPDF fitz)
3. If PNG conversion fails, returns PDF URL for browser fallback

## Common Patterns

### Adding a New Template

1. Create PPTX with `{TOKEN}` placeholders in `data/templates/`
2. Create descriptor JSON with slides array and render_map
3. Add entry to `data/templates/catalog.json`
4. Template will auto-load via TemplateRepository

### Adding a New Input Dataset

1. Create JSON file in `data/inputs/` matching TenantInput schema
2. Add entry to `data/inputs/catalog.json` with id, file, tenant_name, period
3. Dataset will appear in frontend dropdown

### Modifying Slide Content Logic

- **Deterministic mapping**: Edit [data_prep.py](mss_ai_ppt_sample_assets/backend/modules/data_prep.py) `prepare_facts_for_template()`
- **LLM-based generation**: Edit [llm_orchestrator.py](mss_ai_ppt_sample_assets/backend/modules/llm_orchestrator.py) (currently uses mock/deterministic fallback)
- **Token rendering**: Logic is in [ppt_generator.py](mss_ai_ppt_sample_assets/backend/modules/ppt_generator.py) `render()` and `_replace_tokens_in_shape()`

### Frontend State Management

All slide-related state is centralized in `useSlides` store (Zustand):
- `slidespec`, `jobId`, `previews` - core data
- `loading`, `error` - UI state
- `setSlidespec()`, `setPreviews()` - updaters

API calls are made from [App.tsx](frontend/src/App.tsx), results update the store, components re-render automatically.
