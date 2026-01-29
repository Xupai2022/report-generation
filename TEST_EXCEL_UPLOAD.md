# Excel上传端点测试指南

## 端点已集成 ✅

Excel文件上传功能已集成到 `app.py`，位于第314-549行。

### 快速测试

#### 1. 启动服务器

```bash
cd mss_ai_ppt_sample_assets/backend
python app.py
```

服务器启动后访问：http://localhost:8000

应该能看到 `/upload-excel` 出现在端点列表中。

---

#### 2. 使用curl测试上传

**测试正常.xlsx文件：**

```bash
# 准备一个测试Excel文件（按照约定格式）
curl -X POST http://localhost:8000/upload-excel \
  -F "file=@test_report.xlsx"
```

**预期响应：**

```json
{
  "status": "success",
  "session_id": "a3f2c5d8_20250129...",
  "filename": "test_report.xlsx",
  "file_size_mb": 0.45,
  "preview": {
    "customer_name": "测试公司",
    "period": {"start": "2025-12-01", "end": "2025-12-31"},
    "alerts": {"total": 100, "high": 10}
  }
}
```

---

#### 3. 测试安全检查

**测试A：上传.xlsm文件（应被拒绝）**

```bash
curl -X POST http://localhost:8000/upload-excel \
  -F "file=@report.xlsm"
```

**预期响应：** HTTP 400

```json
{
  "detail": {
    "error": "FORBIDDEN_FORMAT",
    "message": "禁止上传 .xlsm 格式文件",
    "detail": "此格式可能包含宏代码，请使用Excel另存为 .xlsx 格式"
  }
}
```

---

**测试B：上传超大文件（> 10MB）**

```bash
# 生成一个11MB的假文件
dd if=/dev/zero of=large.xlsx bs=1M count=11

curl -X POST http://localhost:8000/upload-excel \
  -F "file=@large.xlsx"
```

**预期响应：** HTTP 413

```json
{
  "detail": {
    "error": "FILE_TOO_LARGE",
    "message": "文件大小超过限制 10MB"
  }
}
```

---

#### 4. 完整流程测试（上传 → 生成 → 下载）

```bash
# Step 1: 上传Excel
response=$(curl -X POST http://localhost:8000/upload-excel \
  -F "file=@test_report.xlsx")

# 提取session_id
session_id=$(echo $response | jq -r '.session_id')

echo "Session ID: $session_id"

# Step 2: 生成报告
curl -X POST http://localhost:8000/generate \
  -H "Content-Type: application/json" \
  -d "{
    \"input_id\": \"custom\",
    \"template_id\": \"mss_executive_v2\",
    \"session_id\": \"$session_id\"
  }"

# Step 3: 下载报告
curl -O "http://localhost:8000/download?job_id=${session_id}:mss_executive_v2"
```

---

#### 5. 使用Postman/Insomnia测试

**请求配置：**

- **Method**: POST
- **URL**: `http://localhost:8000/upload-excel`
- **Body**: form-data
  - Key: `file` (type: File)
  - Value: 选择你的 `.xlsx` 文件

**点击Send**，查看响应。

---

## Excel文件格式要求

系统期望的Excel格式（约定）：

| 位置 | 字段 | 说明 |
|------|------|------|
| Sheet名称 | "数据表" | 必须 |
| A2 | 客户名称 | 必填 |
| B2 | 报告开始日期 | 必填 |
| C2 | 报告结束日期 | 必填 |
| D2 | 告警总数 | 可选 |
| E2:H2 | 各级别告警数 | 可选 |
| A5:B10 | 告警分类统计 | 可选 |
| A15:E20 | 事件明细 | 可选 |

---

## 安全机制验证

### ✅ 已实现的安全检查：

1. **文件格式白名单**
   - ✅ 仅接受 `.xlsx`
   - ✅ 拒绝 `.xlsm`, `.xls`, `.xlsb`, `.csv`

2. **MIME类型验证**
   - ✅ 检查 `Content-Type`
   - ✅ 防止文件扩展名伪装

3. **文件大小限制**
   - ✅ 10MB硬限制
   - ✅ 逐块读取，实时检查

4. **会话隔离**
   - ✅ 每个上传独立session_id
   - ✅ 文件存储在隔离目录

5. **只读解析**
   - ✅ `openpyxl` 只读模式
   - ✅ `data_only=True` 不执行公式

---

## 故障排查

### 问题1：导入错误

**错误信息：**
```
ModuleNotFoundError: No module named 'openpyxl'
```

**解决方法：**
```bash
pip install openpyxl
```

---

### 问题2：Excel解析失败

**错误信息：**
```
工作表 '数据表' 不存在
```

**解决方法：**
- 检查Excel文件的Sheet名称是否为"数据表"
- 或修改代码中的工作表名称

---

### 问题3：MIME类型警告

**日志信息：**
```
MIME mismatch: test.xlsx - application/octet-stream
```

**原因：**
浏览器或curl可能未正确识别文件类型。

**解决方法：**
1. 使用标准Excel软件保存文件（不要用WPS等第三方工具）
2. 如果是测试环境，可以临时放宽MIME检查

---

## 下一步

端点已经可用，但需要：

1. ✅ **前端集成**：添加文件上传UI
2. ✅ **Excel模板**：准备标准Excel模板供用户填写
3. ✅ **用户文档**：编写用户操作手册
4. ✅ **监控告警**：添加上传量、失败率监控

---

## 代码位置

- **端点实现**：`app.py` 第314-549行
- **配置常量**：`app.py` 第317-324行
- **Excel解析函数**：`app.py` 第327-408行
- **独立模块参考**：`excel_upload_endpoint.py`（包含集成示例）

---

## 相关文档

- **技术方案**：[docs/excel-to-ppt-workflow.md](docs/excel-to-ppt-workflow.md) (第483-530行)
- **测试验收**：异常场景4（见测试验收文档）
- **API文档**：启动服务器后访问 http://localhost:8000/docs
