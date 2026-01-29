"""Script to enable WebSocket progress updates in app.py

This script adds WebSocket support without replacing the existing concurrency limiter.
Run this to integrate progress updates into your application.
"""

import sys
from pathlib import Path

# Read the current app.py
app_py_path = Path("mss_ai_ppt_sample_assets/backend/app.py")

with open(app_py_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Check if already integrated
if 'WebSocketManager' in content:
    print("❌ WebSocket support already integrated!")
    sys.exit(0)

print("📝 Integrating WebSocket progress updates...")

# 1. Add imports
old_imports = """from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse"""

new_imports = """from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse"""

content = content.replace(old_imports, new_imports)

# 2. Add WebSocketManager import
old_config_import = """from mss_ai_ppt_sample_assets.backend import config
from openai import RateLimitError"""

new_config_import = """from mss_ai_ppt_sample_assets.backend import config
from mss_ai_ppt_sample_assets.backend.websocket_support import WebSocketManager
from openai import RateLimitError"""

content = content.replace(old_config_import, new_config_import)

# 3. Initialize WebSocketManager after service
old_service_init = """service = ReportService()
app.mount("/static/previews", StaticFiles(directory=config.PREVIEWS_DIR), name="previews")"""

new_service_init = """service = ReportService()

# WebSocket Manager for real-time progress updates
ws_manager = WebSocketManager()
logger.info("WebSocket Manager initialized for progress updates")

app.mount("/static/previews", StaticFiles(directory=config.PREVIEWS_DIR), name="previews")"""

content = content.replace(old_service_init, new_service_init)

# 4. Add client_id to GenerateRequest
old_generate_request = """class GenerateRequest(BaseModel):
    input_id: str
    template_id: str
    use_mock: bool = False
    session_id: Optional[str] = None  # Optional: will be auto-generated if not provided"""

new_generate_request = """class GenerateRequest(BaseModel):
    input_id: str
    template_id: str
    use_mock: bool = False
    session_id: Optional[str] = None  # Optional: will be auto-generated if not provided
    client_id: Optional[str] = None  # Optional: WebSocket client ID for progress updates"""

content = content.replace(old_generate_request, new_generate_request)

# 5. Add WebSocket endpoint before @app.get("/")
websocket_endpoint = '''
# WebSocket endpoint for real-time progress updates
@app.websocket("/ws/{client_id}")
async def websocket_endpoint(websocket: WebSocket, client_id: str):
    """WebSocket connection for receiving real-time progress updates.

    Usage from frontend:
        const ws = new WebSocket(`ws://localhost:8000/ws/${clientId}`);
        ws.onmessage = (event) => {
            const data = JSON.parse(event.data);
            console.log(data.type, data.progress, data.message);
        };
    """
    await ws_manager.connect(websocket, client_id)
    try:
        while True:
            # Keep connection alive and handle ping/pong
            data = await websocket.receive_text()
            # Echo back for heartbeat
            if data == "ping":
                await websocket.send_text("pong")
    except WebSocketDisconnect:
        ws_manager.disconnect(client_id)
        logger.info(f"WebSocket client disconnected: {client_id}")


'''

# Find the position to insert (before @app.get("/"))
insert_pos = content.find('@app.get("/")')
if insert_pos != -1:
    content = content[:insert_pos] + websocket_endpoint + content[insert_pos:]

# 6. Register WebSocket session in generate endpoint
old_generate_start = """@app.post("/generate")
async def generate(req: GenerateRequest):
    """Generate a report with OpenAI concurrency limiting.

    This endpoint uses a semaphore to limit concurrent LLM requests,
    preventing API rate limit errors when multiple users generate reports simultaneously.
    """
    logger.info(f"=== Generate Request: input_id={req.input_id}, template_id={req.template_id}, use_mock={req.use_mock}, session_id={req.session_id} ===")"""

new_generate_start = """@app.post("/generate")
async def generate(req: GenerateRequest):
    """Generate a report with OpenAI concurrency limiting.

    This endpoint uses a semaphore to limit concurrent LLM requests,
    preventing API rate limit errors when multiple users generate reports simultaneously.

    If client_id is provided, real-time progress updates will be sent via WebSocket.
    """
    # Generate session_id if not provided
    if not req.session_id:
        import uuid
        req.session_id = f"{uuid.uuid4().hex[:8]}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}"

    # Register WebSocket connection for this session
    if req.client_id:
        ws_manager.register_session(req.session_id, req.client_id)
        logger.info(f"Registered WebSocket: session={req.session_id}, client={req.client_id}")

    logger.info(f"=== Generate Request: input_id={req.input_id}, template_id={req.template_id}, use_mock={req.use_mock}, session_id={req.session_id} ===")"""

content = content.replace(old_generate_start, new_generate_start)

# 7. Add progress notification in generate endpoint
# Find the position after "Acquired LLM slot"
acquired_slot_line = 'logger.info(f"✅ Acquired LLM slot, starting generation...")'
if acquired_slot_line in content:
    # Add progress notification
    progress_notify = '''

            # Send progress update via WebSocket
            if req.client_id:
                await ws_manager.send_progress_update(
                    req.session_id,
                    5,
                    "开始生成报告...",
                    {"template": req.template_id}
                )'''

    content = content.replace(
        acquired_slot_line,
        acquired_slot_line + progress_notify
    )

# 8. Add completion notification
generation_success_line = 'logger.info(f"✓ Generation successful: {result.get(\'job_id\')}")'
if generation_success_line in content:
    completion_notify = '''

            # Send completion notification via WebSocket
            if req.client_id:
                await ws_manager.send_completion(
                    req.session_id,
                    result,
                    success=True
                )'''

    content = content.replace(
        generation_success_line,
        generation_success_line + completion_notify
    )

# Write back
with open(app_py_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("✅ WebSocket support successfully integrated!")
print("")
print("📋 Changes made:")
print("  1. Added WebSocket, BackgroundTasks imports")
print("  2. Initialized WebSocketManager")
print("  3. Added client_id parameter to GenerateRequest")
print("  4. Added /ws/{client_id} WebSocket endpoint")
print("  5. Added progress notifications during generation")
print("  6. Added completion notifications")
print("")
print("🔄 Next steps:")
print("  1. Restart the server:")
print("     python -m mss_ai_ppt_sample_assets.backend.app")
print("")
print("  2. Test WebSocket connection:")
print("     See examples in ADVANCED_FEATURES_GUIDE.md")
