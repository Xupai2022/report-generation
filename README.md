# MSS AI PPT Report Generation System

基于OpenAI的智能PPT报告生成系统，用于自动生成管理安全服务(MSS)月度报告。

## 🚀 功能特性

- ✅ **真实AI生成**：集成OpenAI API，智能生成报告内容
- 🔄 **自动重试机制**：处理网络错误和速率限制
- 📊 **多模板支持**：管理层和技术层两种报告模板
- 🎨 **实时预览**：支持PPT幻灯片预览
- 🔍 **数据校验（可选）**：按需对生成内容做质量检查与提示
- 📝 **审计日志**：完整的操作日志记录

## 📋 系统要求

- Python 3.9+
- Node.js 16+ (前端开发)
- LibreOffice (用于PPT预览生成)
- OpenAI API密钥

## 🔧 安装配置

### 1. 克隆项目

```bash
cd report-generation
```

### 2. 安装Python依赖

```bash
cd mss_ai_ppt_sample_assets/backend
pip install -r requirements.txt
```

### 3. 配置环境变量

复制`.env.example`文件到项目根目录并重命名为`.env`：

```bash
cp .env.example .env
```

编辑`.env`文件，填入您的OpenAI API密钥：

```env
# OpenAI-compatible API Configuration (e.g. Zhipu BigModel)
OPENAI_API_KEY=your-zhipu-api-key-here

# Optional: 自定义OpenAI兼容端点
OPENAI_BASE_URL=https://open.bigmodel.cn/api/paas/v4

# 模型选择
OPENAI_MODEL=glm-4.6

# 启用LLM功能
ENABLE_LLM=true

# 语言设置
DEFAULT_LOCALE=zh-CN
```

### 4. 安装LibreOffice（用于预览）

**Windows:**
```bash
# 下载并安装 LibreOffice
# https://www.libreoffice.org/download/download/

# 或设置环境变量指向安装路径
set LIBREOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe
```

**Linux:**
```bash
sudo apt-get install libreoffice
```

**macOS:**
```bash
brew install --cask libreoffice
```

## 🚀 运行应用

### 启动后端服务

```bash
cd mss_ai_ppt_sample_assets/backend
python app.py
```

服务将在 `http://localhost:8000` 启动

### 启动前端（可选）

```bash
cd frontend
npm install
npm run dev
```

前端将在 `http://localhost:5173` 启动

## 📖 API使用指南

### 1. 查看可用模板

```bash
GET http://localhost:8000/templates
```

### 2. 查看可用输入数据

```bash
GET http://localhost:8000/inputs
```

### 3. 生成报告

```bash
POST http://localhost:8000/generate
Content-Type: application/json

{
  "input_id": "tenant_acme_2025-11",
  "template_id": "mss_management_light_v1"
}
```

**响应示例：**
```json
{
  "job_id": "tenant_acme_2025-11:mss_management_light_v1",
  "report_path": "path/to/report.pptx",
  "warnings": [],
  "slidespec": {...},
  "slidespec_path": "path/to/slidespec.json"
}
```

### 4. 预览报告

```bash
GET http://localhost:8000/api/v1/reports/tenant_acme_2025-11%3Amss_management_light_v1/preview?regenerate_if_missing=true
```

### 5. 下载PPTX

```bash
GET http://localhost:8000/download?job_id=tenant_acme_2025-11:mss_management_light_v1
```

### 6. 重写单页内容

```bash
POST http://localhost:8000/rewrite
Content-Type: application/json

{
  "job_id": "tenant_acme_2025-11:mss_management_light_v1",
  "slide_key": "executive_summary",
  "new_content": {
    "title": "新的标题",
    "bullets": ["更新的要点1", "更新的要点2"]
  }
}
```

## 🔐 安全最佳实践

1. **保护API密钥**：
   - 永远不要将`.env`文件提交到版本控制
   - 在生产环境使用环境变量或密钥管理服务

2. **网络安全**：
   - 在生产环境使用HTTPS
   - 配置适当的CORS策略
   - 添加身份验证和授权机制

3. **速率限制**：
   - OpenAI API有速率限制，系统已内置重试机制
   - 根据需要调整并发请求数量

## 🛠️ 架构说明

### 核心组件

```
mss_ai_ppt_sample_assets/backend/
├── app.py                      # FastAPI应用入口
├── config.py                   # 配置管理（加载.env）
├── services/
│   └── report_service.py      # 报告生成服务
├── modules/
│   ├── llm_orchestrator.py    # OpenAI集成（核心AI逻辑）
│   ├── ppt_generator.py       # PPT生成器
│   ├── preview_generator.py   # 预览生成器
│   ├── (removed)              # 数据验证器（已移除）
│   └── audit_logger.py        # 审计日志
├── models/
│   ├── slidespec.py           # 幻灯片规范模型
│   └── inputs.py              # 输入数据模型
└── data/
    ├── templates/             # PPT模板
    ├── inputs/                # 输入数据
    └── mock_outputs/          # Mock数据（仅用于测试）
```

### AI生成流程

1. **数据准备** (`data_prep.py`): 从输入数据提取关键事实
2. **Prompt构建** (`llm_orchestrator.py`): 根据模板和数据构建提示词
3. **OpenAI调用**: 调用GPT模型生成结构化内容
4. **结果解析**: 解析JSON响应为SlideSpec对象
5. **数据校验（可选）**：可按需对生成内容做提示（默认不做强校验）
6. **PPT渲染** (`ppt_generator.py`): 将内容填充到模板生成最终PPT

### 错误处理

系统实现了多层次的错误处理：

- **网络错误**: 自动重试（指数退避）
- **速率限制**: 智能等待后重试
- **API错误**: 记录日志并降级到确定性生成
- **验证错误**: 返回警告但不阻止生成

## 🧪 测试

### 测试API连通性

```bash
curl http://localhost:8000/health
```

### 测试OpenAI集成

确保`.env`文件中`ENABLE_LLM=true`，然后生成报告：

```bash
curl -X POST http://localhost:8000/generate \
  -H "Content-Type: application/json" \
  -d '{
    "input_id": "tenant_acme_2025-11",
    "template_id": "mss_management_light_v1"
  }'
```

查看日志确认OpenAI调用成功：

```bash
curl http://localhost:8000/logs?limit=50
```

## 📊 监控与日志

系统自动记录以下信息：

- OpenAI API调用状态（成功/失败/重试）
- Token使用量
- 验证警告
- 生成和重写操作

查看日志：

```bash
GET http://localhost:8000/logs?limit=100
```

日志文件位置：`mss_ai_ppt_sample_assets/backend/outputs/logs/`

## 🔄 降级策略

当OpenAI API不可用时，系统会自动降级：

1. **优先级1**: OpenAI生成（`ENABLE_LLM=true`）
2. **优先级2**: Mock数据（如果存在）
3. **优先级3**: 确定性映射（直接使用准备好的数据）

这确保了系统的高可用性。

## 🤝 贡献指南

欢迎提交Issue和Pull Request！

## 📄 许可证

内部使用

## 📞 支持

如有问题请联系开发团队。

---

**注意**: 本系统使用OpenAI API，会产生相应费用。请合理使用并监控API调用量。
