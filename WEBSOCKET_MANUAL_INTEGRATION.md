# WebSocket Progress Updates - 手动集成指南

由于自动集成脚本过于复杂,这里提供简单的手动集成步骤。

## 步骤1: 修改 app.py 导入部分

在 app.py 第1行,将:
```python
from fastapi import FastAPI, HTTPException
```

改为:
```python
from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect
```

在第18行后添加:
```python
from mss_ai_ppt_sample_assets.backend.websocket_support import WebSocketManager
```

## 步骤2: 初始化 WebSocketManager

在第44行 `service = ReportService()` 后添加:
```python
# WebSocket Manager for real-time progress updates
ws_manager = WebSocketManager()
logger.info("✓ WebSocket Manager initialized")
```

## 步骤3: 添加 WebSocket 端点

在 `@app.get("/")` 之前添加:
```python
@app.websocket("/ws/{client_id}")
async def websocket_endpoint(websocket: WebSocket, client_id: str):
    await ws_manager.connect(websocket, client_id)
    try:
        while True:
            data = await websocket.receive_text()
            if data == "ping":
                await websocket.send_text("pong")
    except WebSocketDisconnect:
        ws_manager.disconnect(client_id)
```

## 步骤4: 添加 client_id 参数

在 GenerateRequest 类中添加:
```python
class GenerateRequest(BaseModel):
    input_id: str
    template_id: str
    use_mock: bool = False
    session_id: Optional[str] = None
    client_id: Optional[str] = None  # 新增这一行
```

## 步骤5: 在 generate 端点中注册 session

在 `@app.post("/generate")` 函数开头添加:
```python
# 如果提供了 client_id,注册 WebSocket 连接
if req.client_id and req.session_id:
    ws_manager.register_session(req.session_id, req.client_id)
```

## 步骤6: 发送进度通知

在生成过程中适当位置添加进度通知。

例如在获得槽位后:
```python
# 在 "✅ Acquired LLM slot" 日志后添加
if req.client_id and req.session_id:
    await ws_manager.send_progress_update(
        req.session_id,
        10,
        "开始生成报告...",
        {"template": req.template_id}
    )
```

在生成成功后:
```python
# 在 "✓ Generation successful" 日志后添加
if req.client_id and req.session_id:
    await ws_manager.send_completion(
        req.session_id,
        result,
        success=True
    )
```

## 前端使用示例

```html
<!DOCTYPE html>
<html>
<head>
    <title>WebSocket Progress Demo</title>
</head>
<body>
    <div id="progress"></div>
    <div id="message"></div>

    <script>
        // 生成唯一 client_id
        const clientId = 'client_' + Date.now();

        // 建立 WebSocket 连接
        const ws = new WebSocket(`ws://localhost:8000/ws/${clientId}`);

        ws.onopen = () => {
            console.log('WebSocket connected');

            // 发起生成请求
            fetch('http://localhost:8000/generate', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    input_id: 'tenant_acme_2025-11',
                    template_id: 'mss_executive_v2',
                    client_id: clientId  // 传递 client_id
                })
            })
            .then(r => r.json())
            .then(data => console.log('Generation started:', data));
        };

        ws.onmessage = (event) => {
            const data = JSON.parse(event.data);
            console.log('Progress:', data);

            if (data.type === 'progress') {
                document.getElementById('progress').innerHTML =
                    `Progress: ${data.progress}%`;
                document.getElementById('message').innerHTML =
                    data.message;
            } else if (data.type === 'completed') {
                alert('Generation completed!');
                console.log('Result:', data.result);
            }
        };

        ws.onerror = (error) => {
            console.error('WebSocket error:', error);
        };
    </script>
</body>
</html>
```

## 完成!

按照以上步骤修改后,重启服务器即可使用 WebSocket 进度推送功能。
