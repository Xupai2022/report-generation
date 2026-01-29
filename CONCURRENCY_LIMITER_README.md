# OpenAI 并发限流器实现文档

## 概述

为了防止多用户同时请求时触发 OpenAI API 的速率限制(Rate Limit),我们实现了基于 `asyncio.Semaphore` 的并发限流机制。

## 实现细节

### 核心机制

**位置**: [app.py:37-42](mss_ai_ppt_sample_assets/backend/app.py#L37-L42)

```python
# OpenAI Concurrency Limiter
MAX_CONCURRENT_LLM_REQUESTS = 3  # 最多允许3个并发LLM请求
llm_semaphore = asyncio.Semaphore(MAX_CONCURRENT_LLM_REQUESTS)
```

### 工作原理

1. **Semaphore控制并发数**:
   - 使用 `asyncio.Semaphore` 限制同时进行的 LLM 请求数量
   - 默认值为 3,可根据您的 OpenAI 计划调整

2. **异步请求处理**:
   - `/generate` 端点改为 `async def`,支持异步处理
   - Mock 模式绕过 Semaphore(因为不调用 LLM)
   - 真实 LLM 模式使用 `async with llm_semaphore` 获取锁

3. **队列机制**:
   - 当 3 个 LLM 请求正在处理时,第 4 个请求会等待
   - 一旦某个请求完成释放锁,队列中的下一个请求开始执行

### 代码示例

**位置**: [app.py:94-178](mss_ai_ppt_sample_assets/backend/app.py#L94-L178)

```python
@app.post("/generate")
async def generate(req: GenerateRequest):
    # Mock 模式不需要限流
    if req.use_mock:
        result = service.generate(...)
        return result

    # 真实 LLM 模式:获取 Semaphore
    logger.info("🔄 Waiting for LLM slot...")

    async with llm_semaphore:
        logger.info("✅ Acquired LLM slot")

        # 在线程池中执行同步代码
        result = await loop.run_in_executor(
            None,
            lambda: service.generate(...)
        )

        logger.info("🔓 Released LLM slot")
        return result
```

## 监控和调试

### Health Check API

**端点**: `GET /health`

**响应示例**:
```json
{
  "status": "ok",
  "llm_concurrency": {
    "max_concurrent_requests": 3,
    "available_slots": 1,
    "active_requests": 2,
    "is_saturated": false
  }
}
```

**字段说明**:
- `max_concurrent_requests`: 最大并发数
- `available_slots`: 当前可用槽位数
- `active_requests`: 正在进行的请求数
- `is_saturated`: 是否已满载(无可用槽位)

### 日志输出

系统会记录详细的并发状态日志:

```
🔄 Waiting for LLM slot (current waiting: 2)
✅ Acquired LLM slot, starting generation...
✓ Generation successful: a3f2c5d8_20250129:mss_executive_v2
🔓 Released LLM slot
```

## 测试

### 安装测试依赖

```bash
pip install aiohttp
```

### 运行并发测试

```bash
# 确保后端服务正在运行
cd mss_ai_ppt_sample_assets/backend
python app.py

# 在新终端运行测试
cd /path/to/project
python test_concurrency_limiter.py
```

### 测试场景

测试脚本包含 3 个测试场景:

1. **Mock 模式** - 验证基线性能(无限流)
2. **适度负载** - 5 个并发请求
3. **高负载** - 10 个并发请求(超过限制)

**预期结果**:
- Mock 模式: 所有请求快速完成,无排队
- 真实 LLM 模式: 每次最多 3 个并发,其余请求排队等待
- 高负载: 后续请求等待时间明显增加

## 调优建议

### 调整并发数

根据您的 OpenAI 计划调整 `MAX_CONCURRENT_LLM_REQUESTS`:

| OpenAI 计划 | 建议并发数 | 说明 |
|-------------|-----------|------|
| Free Tier | 1-2 | 限制严格,保守设置 |
| Pay-as-you-go | 3-5 | 标准设置 |
| Enterprise | 10+ | 根据实际限制调整 |

**修改位置**: [app.py:40](mss_ai_ppt_sample_assets/backend/app.py#L40)

```python
MAX_CONCURRENT_LLM_REQUESTS = 5  # 调整为 5
```

### 超时处理

如果请求长时间排队,可以添加超时机制:

```python
async with asyncio.timeout(300):  # 5 分钟超时
    async with llm_semaphore:
        result = await loop.run_in_executor(...)
```

## 并发安全性总结

✅ **已实现的安全机制**:
1. Session 隔离 - 不同用户文件独立
2. 文件锁 (FileLock) - 防止写入冲突
3. **OpenAI 并发限流** - 防止 API 限流

⚠️ **注意事项**:
- 限流器只对真实 LLM 请求生效
- Mock 模式绕过限流(用于测试)
- 调整并发数需重启服务

## 性能影响

**优势**:
- 防止 OpenAI API 返回 429 错误
- 平滑处理高并发场景
- 提供可见的队列状态

**劣势**:
- 超过限制的请求需等待
- 可能增加用户等待时间

**最佳实践**:
1. 根据实际 API 配额设置并发数
2. 监控 `/health` 端点了解负载
3. 考虑添加前端排队提示

## 故障排查

### 问题: 所有请求都很慢

**可能原因**: 并发数设置过低

**解决方案**: 增加 `MAX_CONCURRENT_LLM_REQUESTS`

### 问题: 仍然触发 Rate Limit

**可能原因**:
1. 并发数设置过高
2. 其他服务也在调用同一 API key

**解决方案**:
1. 降低并发数
2. 使用独立 API key

### 问题: 请求卡住不动

**可能原因**: Semaphore 死锁

**解决方案**: 重启服务,检查 finally 块是否正确释放锁

## 未来优化

1. **动态调整**: 根据 API 响应自动调整并发数
2. **优先级队列**: VIP 用户优先处理
3. **降级策略**: API 限流时自动切换到 Mock 模式
4. **分布式限流**: 多服务器共享限流配额

---

**实现作者**: Claude + User
**实现日期**: 2025-01-29
**版本**: 1.0
