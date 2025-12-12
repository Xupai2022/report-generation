# Windows 环境部署指南

本文档详细说明如何在新的 Windows 机器上从零开始部署此项目。

## 📋 前置要求

### 1. 安装 Python 3.9+

**下载地址**: https://www.python.org/downloads/

安装时**务必勾选**"Add Python to PATH"选项。

验证安装：
```cmd
python --version
pip --version
```

### 2. 安装 LibreOffice

**下载地址**: https://www.libreoffice.org/download/download/

**推荐版本**: LibreOffice 7.6+ (最新稳定版)

**安装路径**:
- 默认路径: `C:\Program Files\LibreOffice\program\soffice.exe`
- 或自定义路径 (需要后续配置环境变量)

验证安装：
```cmd
"C:\Program Files\LibreOffice\program\soffice.exe" --version
```

如果显示版本信息则安装成功。

### 3. 安装 Git (可选，用于克隆项目)

**下载地址**: https://git-scm.com/download/win

或者直接从 GitHub/其他源下载项目压缩包解压。

---

## 🚀 部署步骤

### Step 1: 获取项目代码

**方法 A - 使用 Git**:
```cmd
cd C:\Users\YourName\Desktop
git clone <your-repo-url> report-generation
cd report-generation
```

**方法 B - 下载压缩包**:
解压到目标目录，例如 `C:\Users\YourName\Desktop\report-generation`

### Step 2: 安装 Python 依赖

```cmd
cd mss_ai_ppt_sample_assets\backend
pip install -r requirements.txt
```

预期输出：
```
Successfully installed fastapi-0.110.0 uvicorn-0.23.0 pydantic-2.0.0 ...
```

### Step 3: 配置 LibreOffice 路径 (如果非默认安装)

如果 LibreOffice 安装在非默认路径，需要设置环境变量。

**临时设置 (仅当前命令行窗口有效)**:
```cmd
set LIBREOFFICE_PATH=D:\MyPrograms\LibreOffice\program\soffice.exe
```

**永久设置 (推荐)**:
1. 右键"此电脑" → "属性"
2. 点击"高级系统设置"
3. 点击"环境变量"
4. 在"系统变量"区域点击"新建"
5. 变量名: `LIBREOFFICE_PATH`
6. 变量值: `C:\Program Files\LibreOffice\program\soffice.exe` (根据实际路径修改)
7. 点击"确定"保存

**验证配置**:
```cmd
# 重新打开命令行窗口
echo %LIBREOFFICE_PATH%
```

### Step 4: 配置 OpenAI API (可选)

如果需要使用 LLM 功能，创建 `.env` 文件：

```cmd
cd C:\Users\YourName\Desktop\report-generation
notepad .env
```

在记事本中输入以下内容：
```env
# OpenAI API Configuration
OPENAI_API_KEY=sk-your-openai-api-key-here
OPENAI_MODEL=gpt-4o-mini
ENABLE_LLM=true
DEFAULT_LOCALE=zh-CN
```

保存文件。

**注意**: 如果不需要 LLM 功能，可以跳过此步骤，系统会使用 mock 数据。

---

## ✅ 验证部署

### 1. 测试 LibreOffice 转换功能

```cmd
cd mss_ai_ppt_sample_assets\backend
python -c "from modules.preview_generator import PPTPreviewGenerator; print('Preview generator OK')"
```

如果没有报错，说明 LibreOffice 配置正确。

### 2. 启动后端服务

```cmd
cd mss_ai_ppt_sample_assets\backend
python app.py
```

预期输出：
```
INFO:     Started server process [12345]
INFO:     Waiting for application startup.
INFO:     Application startup complete.
INFO:     Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)
```

### 3. 测试 API

打开浏览器访问: http://localhost:8000/docs

你应该能看到 FastAPI 自动生成的 API 文档。

### 4. 测试预览生成 (完整流程)

**使用 PowerShell 或 Git Bash**:
```powershell
# 1. 查看可用模板
curl http://localhost:8000/templates

# 2. 查看可用输入
curl http://localhost:8000/inputs

# 3. 生成报告
curl -X POST http://localhost:8000/generate `
  -H "Content-Type: application/json" `
  -d '{\"input_id\":\"tenant_acme_2025-11\",\"template_id\":\"mss_management_light_v1\"}'

# 4. 生成预览 (使用上一步返回的 job_id)
curl "http://localhost:8000/preview?job_id=tenant_acme_2025-11:mss_management_light_v1"
```

**使用浏览器**:
直接访问 http://localhost:8000/docs，在 Swagger UI 中测试各个接口。

---

## 🔧 常见问题排查

### 问题 1: `LibreOffice not found`

**症状**:
```
PreviewGenerationError: LibreOffice (`soffice`) is required...
```

**解决方案**:
1. 检查 LibreOffice 是否已安装:
   ```cmd
   "C:\Program Files\LibreOffice\program\soffice.exe" --version
   ```

2. 如果路径不同，设置环境变量:
   ```cmd
   set LIBREOFFICE_PATH=你的实际路径\soffice.exe
   ```

3. 确认路径中没有中文或特殊字符

### 问题 2: `ModuleNotFoundError: No module named 'fitz'`

**症状**:
```
PyMuPDF (pymupdf) is required for PDF to PNG conversion.
```

**解决方案**:
```cmd
pip install pymupdf
```

### 问题 3: 预览图片生成失败

**症状**:
API 返回错误或图片为空

**排查步骤**:

1. 检查 `outputs/previews/` 目录是否存在:
   ```cmd
   dir mss_ai_ppt_sample_assets\backend\outputs\previews
   ```

2. 手动测试 LibreOffice 转换:
   ```cmd
   cd mss_ai_ppt_sample_assets\backend\outputs\reports
   "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf test.pptx
   ```

3. 查看审计日志:
   ```cmd
   type mss_ai_ppt_sample_assets\backend\outputs\logs\audit.jsonl
   ```

### 问题 4: 端口 8000 被占用

**症状**:
```
OSError: [WinError 10048] Only one usage of each socket address...
```

**解决方案 A - 更换端口**:
```cmd
python -m uvicorn mss_ai_ppt_sample_assets.backend.app:app --reload --port 8001
```

**解决方案 B - 关闭占用端口的进程**:
```cmd
netstat -ano | findstr :8000
taskkill /PID <进程ID> /F
```

### 问题 5: 中文乱码

**症状**:
生成的 PPT 或预览中中文显示为乱码

**解决方案**:
1. 确保 Python 脚本使用 UTF-8 编码
2. 检查 `.env` 文件:
   ```env
   DEFAULT_LOCALE=zh-CN
   ```

3. 确保 PowerPoint 模板中使用了支持中文的字体 (如微软雅黑)

---

## 🎯 生产环境建议

### 1. 使用虚拟环境

```cmd
cd report-generation
python -m venv venv
venv\Scripts\activate
cd mss_ai_ppt_sample_assets\backend
pip install -r requirements.txt
```

### 2. 后台运行服务

**使用 Windows 任务计划程序**:
1. 打开"任务计划程序"
2. 创建基本任务
3. 触发器: 系统启动时
4. 操作: 启动程序
   - 程序: `C:\Users\YourName\Desktop\report-generation\venv\Scripts\python.exe`
   - 参数: `mss_ai_ppt_sample_assets\backend\app.py`
   - 起始于: `C:\Users\YourName\Desktop\report-generation`

**或使用 NSSM (Non-Sucking Service Manager)**:
```cmd
# 下载 NSSM: https://nssm.cc/download
nssm install ReportGenerationAPI "C:\...\python.exe" "C:\...\app.py"
nssm start ReportGenerationAPI
```

### 3. 防火墙配置

允许端口 8000 的入站连接:
```cmd
netsh advfirewall firewall add rule name="Report Generation API" dir=in action=allow protocol=TCP localport=8000
```

---

## 📂 目录权限检查

确保以下目录有写入权限:

```
mss_ai_ppt_sample_assets\backend\
├── outputs\
│   ├── reports\      # 生成的 PPTX 文件
│   ├── previews\     # 预览图片
│   ├── slidespecs\   # 中间 JSON 文件
│   └── logs\         # 审计日志
```

如果权限不足:
```cmd
icacls outputs /grant Users:(OI)(CI)F /T
```

---

## 🔄 更新部署

当代码更新后:

```cmd
cd report-generation
git pull  # 或重新解压新版本

cd mss_ai_ppt_sample_assets\backend
pip install -r requirements.txt --upgrade

# 重启服务
```

---

## 📞 支持

遇到问题请检查:
1. Python 版本 >= 3.9
2. LibreOffice 已正确安装
3. 所有 pip 依赖已安装
4. 环境变量配置正确
5. 文件夹权限充足

详细日志位置: `mss_ai_ppt_sample_assets\backend\outputs\logs\audit.jsonl`
