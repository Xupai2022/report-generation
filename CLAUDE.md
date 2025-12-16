# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MSS AI PPT is a report generation platform that creates PowerPoint presentations from structured security data. The system uses V2 AI-driven templates where:
- Raw `TenantInput` data goes directly to OpenAI LLM
- AI generates content based on `{{TOKEN}}` placeholder instructions in template descriptors
- Generated content is validated against key numerical fields

**Architecture**: FastAPI backend with static HTML frontend (no build step required)

## Development Commands

### Backend (FastAPI)

```bash
# Run backend server
cd mss_ai_ppt_sample_assets/backend
python app.py
# Server runs at http://localhost:8000

# Or with uvicorn for hot-reload
python -m uvicorn mss_ai_ppt_sample_assets.backend.app:app --reload --port 8000
```

### Dependencies

```bash
cd mss_ai_ppt_sample_assets/backend
pip install -r requirements.txt
```

**External requirement**: LibreOffice must be installed for preview generation (PPTX → PDF → PNG pipeline).

## Code Architecture

### Backend Pipeline (mss_ai_ppt_sample_assets/backend/)

The V2 pipeline uses AI-driven content generation:

1. **ReportService** ([services/report_service.py](mss_ai_ppt_sample_assets/backend/services/report_service.py))
   - Entry point: `generate()`, `rewrite()`, `preview()`
   - Only supports V2 templates (raises error for V1)
   - Coordinates: LLM generation → validation → PPT rendering → preview

2. **Core Modules** (modules/):
   - **llm_orchestrator.py** - `LLMOrchestratorV2` generates SlideSpecV2 via OpenAI
     - Extracts data placeholders (non-AI) from TenantInput using dotted paths
     - Generates AI placeholders with smart batching to avoid timeouts
     - Supports chart and table data extraction
   - **ppt_generator.py** - `PPTGeneratorV2` renders SlideSpecV2 to PPTX
     - Replaces `{{TOKEN}}` placeholders in template
     - Renders native charts (bar/pie) and tables via python-pptx
   - **preview_generator.py** - Converts PPTX → PDF (LibreOffice) → PNG (PyMuPDF)
   - **template_loader.py** - Loads template catalog and descriptors
   - **validator.py** - `ValidatorV2` validates key numbers match input data
   - **audit_logger.py** - Writes structured events to `outputs/logs/audit.jsonl`

3. **Data Models** (models/):
   - **inputs.py** - `TenantInput` wraps raw JSON with `.get()` accessor
   - **slidespec.py** - `SlideSpecV2` with `slide_no`, `slide_key`, `placeholders` dict
   - **templates.py** - `TemplateDescriptorV2` with `PlaceholderDefinition` for each token
   - **audit.py** - `AuditEvent` schema

4. **Configuration** ([config.py](mss_ai_ppt_sample_assets/backend/config.py)):
   - Loads `.env` from project root
   - `Settings` class with `enable_llm`, `openai_api_key`, `openai_model`, etc.

### Key Data Flow

```
POST /generate (input_id, template_id)
  → ReportService.generate()
    → LLMOrchestratorV2.generate_slidespec_v2()
      → Extract data placeholders (source paths, charts, tables)
      → Generate AI placeholders via OpenAI (with smart batching)
    → ValidatorV2.validate_key_numbers()
    → PPTGeneratorV2.render() → outputs/reports/{input_id}_{template_id}.pptx
    → Save SlideSpecV2 → outputs/slidespecs/{input_id}_{template_id}.json

GET /preview?job_id={input_id}:{template_id}
  → PPTPreviewGenerator.to_images()
    → LibreOffice: PPTX → PDF
    → PyMuPDF: PDF → PNG images
  → Returns URLs: /static/previews/{job_id}/slide*.png
```

### Template System (V2)

Templates consist of two files in `mss_ai_ppt_sample_assets/backend/data/templates/`:
- **PPTX file** - PowerPoint with `{{TOKEN}}` placeholders in text
- **descriptor JSON** - Defines slides with placeholder definitions

**PlaceholderDefinition** key fields:
- `token`: Placeholder name (e.g., "HEADLINE")
- `type`: "text", "paragraph", "bullet_list", "bar_chart", "pie_chart", "native_table"
- `ai_generate`: true = AI generates content, false = extract from data
- `source`: Dotted path for data extraction (e.g., "alerts.total")
- `ai_instruction`: Prompt for AI generation
- `chart_config` / `table_config`: Configuration for charts/tables

Current templates:
- `mss_executive_v2` - 8 slides (management audience)
- `mss_technical_v2` - 10 slides (technical audience)

### Environment Variables

```bash
# .env file in project root
OPENAI_API_KEY=sk-...          # Required if ENABLE_LLM=true
OPENAI_BASE_URL=https://...    # Optional: custom endpoint
OPENAI_MODEL=gpt-4o-mini       # Default model
ENABLE_LLM=true                # Enable real LLM (default: false, uses mock)
DEFAULT_LOCALE=zh-CN           # Default locale
```

### Important File Locations

```
mss_ai_ppt_sample_assets/backend/
├── data/
│   ├── inputs/           # Input JSON files + catalog.json
│   ├── templates/        # PPTX templates + descriptor JSON + catalog.json
│   └── mock_outputs/     # Pre-generated SlideSpecs for testing
├── outputs/
│   ├── reports/          # Generated PPTX: {input_id}_{template_id}.pptx
│   ├── slidespecs/       # Saved JSON: {input_id}_{template_id}.json
│   ├── previews/         # PNG images: {job_id}/slide*.png
│   └── logs/             # audit.jsonl
└── frontend/             # Static HTML UI (served at /ui)
```

## Common Patterns

### Adding a New V2 Template

1. Create PPTX with `{{TOKEN}}` placeholders in `data/templates/`
2. Create descriptor JSON defining each placeholder:
   ```json
   {
     "template_id": "my_template_v2",
     "slides": [{
       "slide_no": 1,
       "slide_key": "cover",
       "placeholders": [
         {"token": "TITLE", "type": "text", "ai_generate": true, "ai_instruction": "..."},
         {"token": "DATE", "type": "text", "ai_generate": false, "source": "period.start"}
       ]
     }]
   }
   ```
3. Add entry to `data/templates/catalog.json`

### Adding a New Input Dataset

1. Create JSON file in `data/inputs/` with security data (alerts, incidents, vulnerabilities, etc.)
2. Add entry to `data/inputs/catalog.json`:
   ```json
   {"id": "tenant_acme_2025-12", "file": "tenant_acme_2025-12.json", ...}
   ```

### Modifying AI Generation

- **System/user prompts**: `LLMOrchestratorV2._build_system_prompt()` and `_build_user_prompt()` in [llm_orchestrator.py](mss_ai_ppt_sample_assets/backend/modules/llm_orchestrator.py)
- **Data extraction**: `_extract_data_placeholders()`, `_extract_chart_data()`, `_extract_table_data()`
- **Batching logic**: `_get_smart_slide_batches()` splits slides to avoid API timeouts
- **Fallback content**: `_fill_ai_placeholders_with_fallback()` when LLM is unavailable

### Token Rendering

- Text tokens: `PPTGeneratorV2._replace_tokens_in_shape()` replaces `{{TOKEN}}` in text frames
- Charts: `_render_bar_chart()`, `_render_pie_chart()` add native PowerPoint charts
- Tables: `_render_native_table()` adds PowerPoint tables with headers and data rows
