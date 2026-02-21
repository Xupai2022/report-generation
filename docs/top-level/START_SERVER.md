# 服务器启动指南

## 🚨 重要: 正确的启动方式

### ❌ 错误方式
```bash
cd mss_ai_ppt_sample_assets/backend
python app.py  # 会导致 ModuleNotFoundError
```

### ✅ 正确方式

**方法1: 从项目根目录运行(推荐)**
```bash
cd /path/to/report-generation
python -m mss_ai_ppt_sample_assets.backend.app
```

**方法2: 设置PYTHONPATH**
```bash
# Windows
set PYTHONPATH=C:\Users\User\Desktop\report-generation
cd mss_ai_ppt_sample_assets\backend
python app.py

# Linux/Mac
export PYTHONPATH=/path/to/report-generation
cd mss_ai_ppt_sample_assets/backend
python app.py
```

## 检查服务状态

### 1. 检查端口监听
```bash
# Windows
netstat -ano | findstr :8000

# Linux/Mac
lsof -i :8000
```

应该看到:
```
TCP    0.0.0.0:8000           0.0.0.0:0              LISTENING       12345
```

### 2. 访问健康检查
```bash
curl http://localhost:8000/health
```

应该返回:
```json
{
  "status": "ok",
  "llm_concurrency": {
    "max_concurrent_requests": 3,
    "available_slots": 3,
    "active_requests": 0,
    "is_saturated": false
  }
}
```

### 3. 访问Web UI
打开浏览器: http://localhost:8000/ui/

## 运行测试

```bash
# 确保服务已启动
python tests/test_concurrency_limiter.py
```

## 常见问题

### Q: ModuleNotFoundError: No module named 'mss_ai_ppt_sample_assets'

**原因**: 没有从项目根目录运行

**解决**: 使用方法1或方法2启动

### Q: 端口8000已被占用

**检查占用进程**:
```bash
# Windows
netstat -ano | findstr :8000
taskkill /F /PID <进程ID>

# Linux/Mac
lsof -i :8000
kill -9 <PID>
```

### Q: 测试脚本无法连接

**检查端口配置**:
- 服务器默认端口: 8000
- 测试脚本配置: [test_concurrency_limiter.py:17](../../tests/test_concurrency_limiter.py#L17)

确保两者一致!

## 启动日志示例

成功启动应该看到:
```
INFO:     Started server process [12345]
INFO:     Waiting for application startup.
Initializing MSS AI PPT Backend...
LLM Enabled: True
OpenAI Model: gpt-4o-mini
LLM Concurrency Limiter: max 3 concurrent requests
INFO:     Application startup complete.
INFO:     Uvicorn running on http://0.0.0.0:8000 (Press CTRL+C to quit)
```
