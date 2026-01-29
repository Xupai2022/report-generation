"""Simple patch to add WebSocket support to app.py"""

app_file = "mss_ai_ppt_sample_assets/backend/app.py"

# Read current content
with open(app_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Create new content list
new_lines = []
added_ws_manager = False
added_websocket_endpoint = False

for i, line in enumerate(lines):
    # 1. Update FastAPI import
    if line.startswith('from fastapi import FastAPI, HTTPException'):
        new_lines.append('from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect\n')
        continue

    # 2. Add WebSocketManager import after config import
    if 'from mss_ai_ppt_sample_assets.backend import config' in line:
        new_lines.append(line)
        new_lines.append('from mss_ai_ppt_sample_assets.backend.websocket_support import WebSocketManager\n')
        continue

    # 3. Add ws_manager after service initialization
    if line.startswith('service = ReportService()') and not added_ws_manager:
        new_lines.append(line)
        new_lines.append('\n# WebSocket Manager for real-time progress updates\n')
        new_lines.append('ws_manager = WebSocketManager()\n')
        new_lines.append('logger.info("WebSocket Manager initialized")\n')
        added_ws_manager = True
        continue

    # 4. Add client_id to GenerateRequest
    if 'session_id: Optional[str] = None  # Optional: will be auto-generated' in line:
        new_lines.append(line)
        new_lines.append('    client_id: Optional[str] = None  # Optional: WebSocket client ID\n')
        continue

    # 5. Add WebSocket endpoint before root endpoint
    if line.strip() == '@app.get("/")' and not added_websocket_endpoint:
        # Add WebSocket endpoint
        new_lines.append('@app.websocket("/ws/{client_id}")\n')
        new_lines.append('async def websocket_endpoint(websocket: WebSocket, client_id: str):\n')
        new_lines.append('    await ws_manager.connect(websocket, client_id)\n')
        new_lines.append('    try:\n')
        new_lines.append('        while True:\n')
        new_lines.append('            data = await websocket.receive_text()\n')
        new_lines.append('            if data == "ping":\n')
        new_lines.append('                await websocket.send_text("pong")\n')
        new_lines.append('    except WebSocketDisconnect:\n')
        new_lines.append('        ws_manager.disconnect(client_id)\n')
        new_lines.append('\n\n')
        added_websocket_endpoint = True
        new_lines.append(line)
        continue

    # 6. Register WebSocket session at start of generate endpoint
    if '    logger.info(f"=== Generate Request: input_id={req.input_id}' in line:
        new_lines.append('    # Register WebSocket connection if client_id provided\n')
        new_lines.append('    if req.client_id and req.session_id:\n')
        new_lines.append('        ws_manager.register_session(req.session_id, req.client_id)\n')
        new_lines.append('\n')
        new_lines.append(line)
        continue

    # Default: keep the line
    new_lines.append(line)

# Write back
with open(app_file, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("✓ WebSocket support added successfully!")
print("")
print("Changes made:")
print("  1. Added WebSocket, WebSocketDisconnect imports")
print("  2. Added WebSocketManager import and initialization")
print("  3. Added client_id to GenerateRequest")
print("  4. Added /ws/{client_id} WebSocket endpoint")
print("  5. Added session registration in /generate")
print("")
print("To send progress updates, add this in your generation code:")
print("  await ws_manager.send_progress_update(session_id, progress, message)")
print("")
print("Restart server to apply changes!")
