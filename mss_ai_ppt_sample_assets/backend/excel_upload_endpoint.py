"""
Excel文件上传端点（独立模块，暂未集成到主应用）

功能：
- 接收用户上传的Excel文件
- 验证文件格式和大小
- 解析Excel数据为JSON
- 返回session_id供后续生成报告使用

安全特性：
- 仅接受 .xlsx 格式（无宏Excel）
- 文件大小限制 10MB
- MIME类型双重验证
- 会话隔离存储

集成方法：
1. 将 upload_excel() 函数复制到 app.py
2. 添加 extract_excel_data() 导入
3. 重启服务器测试
"""

from fastapi import UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from pathlib import Path
import json
import logging
from typing import Optional
import openpyxl

logger = logging.getLogger(__name__)

# 配置常量
ALLOWED_EXTENSIONS = {'.xlsx'}  # 只允许无宏Excel
FORBIDDEN_EXTENSIONS = {'.xlsm', '.xls', '.xlsb', '.csv'}  # 明确禁止
ALLOWED_MIME_TYPES = {
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/zip',  # .xlsx本质是zip，部分系统识别为此
}
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB
CHUNK_SIZE = 8192  # 8KB chunks


def extract_excel_data(excel_path: str) -> dict:
    """从Excel文件提取约定格式的数据

    约定的Excel格式：
    - Sheet名称: "数据表"
    - A2: 客户名称
    - B2:C2: 报告周期（开始、结束）
    - D2: 告警总数
    - E2:H2: 告警级别数量（高、中、低、信息）
    - A5:B10: 告警分类数据（类别、数量）
    - A15:E20: 事件明细（ID、类型、严重程度、状态、描述）

    Args:
        excel_path: Excel文件路径

    Returns:
        结构化的数据字典

    Raises:
        ValueError: 必填字段缺失或数据格式错误
    """
    try:
        # 只读模式打开，data_only=True 只读取值（忽略公式）
        wb = openpyxl.load_workbook(
            excel_path,
            read_only=True,
            data_only=True
        )

        # 检查工作表是否存在
        if '数据表' not in wb.sheetnames:
            raise ValueError(
                f"工作表 '数据表' 不存在。找到的工作表: {', '.join(wb.sheetnames)}"
            )

        ws = wb['数据表']

        # 提取基础信息（必填字段验证）
        customer_name = ws['A2'].value
        if not customer_name:
            raise ValueError("必填字段 'customer_name' (A2) 为空")

        period_start = ws['B2'].value
        if not period_start:
            raise ValueError("必填字段 'period.start' (B2) 为空")

        period_end = ws['C2'].value
        if not period_end:
            raise ValueError("必填字段 'period.end' (C2) 为空")

        # 构建数据结构
        data = {
            "customer_name": str(customer_name).strip(),
            "period": {
                "start": str(period_start),
                "end": str(period_end)
            },

            # 提取统计数据
            "alerts": {
                "total": ws['D2'].value or 0,
                "by_severity": {
                    "high": ws['E2'].value or 0,
                    "medium": ws['F2'].value or 0,
                    "low": ws['G2'].value or 0,
                    "info": ws['H2'].value or 0
                }
            },

            # 提取分类数据（用于柱状图）
            "top_alert_categories": [],

            # 提取明细数据（用于表格）
            "top_incidents": []
        }

        # 提取告警分类数据（A5:B10）
        for row in range(5, 11):
            category = ws[f'A{row}'].value
            count = ws[f'B{row}'].value
            if category and count is not None:  # 允许count=0
                data["top_alert_categories"].append({
                    "category": str(category).strip(),
                    "count": int(count) if isinstance(count, (int, float)) else 0
                })

        # 提取事件明细（A15:E20）
        for row in range(15, 21):
            incident_id = ws[f'A{row}'].value
            if incident_id:  # 至少要有ID
                data["top_incidents"].append({
                    "id": str(incident_id).strip(),
                    "type": str(ws[f'B{row}'].value or "").strip(),
                    "severity": str(ws[f'C{row}'].value or "").strip(),
                    "status": str(ws[f'D{row}'].value or "").strip(),
                    "description": str(ws[f'E{row}'].value or "").strip()
                })

        wb.close()

        logger.info(
            f"✅ Excel parsed successfully: {len(data['top_alert_categories'])} categories, "
            f"{len(data['top_incidents'])} incidents"
        )

        return data

    except openpyxl.utils.exceptions.InvalidFileException as e:
        raise ValueError(f"无效的Excel文件格式: {str(e)}")
    except Exception as e:
        raise ValueError(f"Excel解析失败: {str(e)}")


def validate_excel_format(
    filename: str,
    content_type: str,
    file_size: int
) -> None:
    """验证Excel文件格式和大小

    Args:
        filename: 文件名
        content_type: MIME类型
        file_size: 文件大小（字节）

    Raises:
        HTTPException: 验证失败
    """
    # 1. 文件扩展名检查
    file_ext = Path(filename).suffix.lower()

    if file_ext not in ALLOWED_EXTENSIONS:
        # 如果是明确禁止的格式，给出详细说明
        if file_ext in FORBIDDEN_EXTENSIONS:
            format_tips = {
                '.xlsm': '此格式可能包含宏代码，请使用Excel另存为 .xlsx 格式',
                '.xls': '旧版Excel格式，请使用Excel 2007+另存为 .xlsx 格式',
                '.xlsb': '二进制Excel格式，请另存为 .xlsx 格式',
                '.csv': 'CSV格式不支持，请使用Excel打开并另存为 .xlsx 格式'
            }
            detail_message = format_tips.get(file_ext, '请转换为 .xlsx 格式')

            raise HTTPException(
                status_code=400,
                detail={
                    "error": "FORBIDDEN_FORMAT",
                    "message": f"禁止上传 {file_ext} 格式文件",
                    "detail": detail_message,
                    "allowed_formats": list(ALLOWED_EXTENSIONS)
                }
            )
        else:
            raise HTTPException(
                status_code=400,
                detail={
                    "error": "UNSUPPORTED_FORMAT",
                    "message": f"不支持的文件格式 {file_ext}，仅允许 .xlsx",
                    "allowed_formats": list(ALLOWED_EXTENSIONS)
                }
            )

    # 2. MIME类型检查（防止文件伪装）
    if content_type not in ALLOWED_MIME_TYPES:
        logger.warning(
            f"MIME type mismatch: filename={filename}, "
            f"content_type={content_type}, expected={ALLOWED_MIME_TYPES}"
        )
        raise HTTPException(
            status_code=400,
            detail={
                "error": "MIME_TYPE_MISMATCH",
                "message": "文件类型验证失败，文件可能被伪装",
                "detail": f"检测到 {content_type}，但期望 .xlsx 格式",
                "suggestion": "请确保文件是真正的 Excel 2007+ 格式"
            }
        )

    # 3. 文件大小检查
    if file_size > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=413,
            detail={
                "error": "FILE_TOO_LARGE",
                "message": f"文件大小 {file_size / 1024 / 1024:.2f}MB 超过限制 {MAX_FILE_SIZE / 1024 / 1024}MB",
                "suggestion": "请删除Excel中的图片、图表等多余内容后重新上传"
            }
        )


async def upload_excel(
    file: UploadFile = File(...),
    session_manager = None,  # 注入session_manager
    sessions_dir: Path = None  # 注入sessions目录
) -> JSONResponse:
    """Excel文件上传端点

    功能：
    1. 验证文件格式（.xlsx）和大小（< 10MB）
    2. 保存文件到会话隔离目录
    3. 解析Excel数据为JSON
    4. 返回session_id和数据预览

    Args:
        file: 上传的文件
        session_manager: 会话管理器（需要从外部注入）
        sessions_dir: 会话存储目录（需要从外部注入）

    Returns:
        JSONResponse包含：
        - status: "success"
        - session_id: 唯一会话ID
        - filename: 原始文件名
        - file_size_mb: 文件大小
        - preview: 数据预览（客户名称、周期、告警总数等）

    Raises:
        HTTPException:
            - 400: 文件格式/数据验证失败
            - 413: 文件过大
            - 500: 服务器内部错误
    """
    logger.info(f"📥 Upload request: filename={file.filename}, content_type={file.content_type}")

    # 预先进行轻量级验证（不读取整个文件）
    validate_excel_format(
        filename=file.filename,
        content_type=file.content_type,
        file_size=0  # 先不检查大小，边读边检查
    )

    # 生成会话ID
    if session_manager:
        session_id = session_manager.generate_session_id()
    else:
        # 如果没有注入session_manager，使用简单的UUID生成
        import uuid
        from datetime import datetime
        uuid_part = uuid.uuid4().hex[:8]
        timestamp_part = datetime.now().strftime('%Y%m%d%H%M%S%f')
        session_id = f"{uuid_part}_{timestamp_part}"

    # 创建会话目录
    if sessions_dir is None:
        # 默认路径（需要根据实际项目调整）
        sessions_dir = Path(__file__).parent / "outputs" / "sessions"

    session_dir = sessions_dir / session_id
    session_dir.mkdir(parents=True, exist_ok=True)

    file_path = session_dir / "uploaded.xlsx"
    file_size = 0

    try:
        # 逐块读取并保存（边读边检查大小）
        with file_path.open("wb") as f:
            while chunk := await file.read(CHUNK_SIZE):
                file_size += len(chunk)

                # 实时检查文件大小
                if file_size > MAX_FILE_SIZE:
                    # 立即停止读取并删除文件
                    f.close()
                    file_path.unlink()
                    raise HTTPException(
                        status_code=413,
                        detail={
                            "error": "FILE_TOO_LARGE",
                            "message": f"文件大小超过限制 {MAX_FILE_SIZE / 1024 / 1024}MB",
                            "actual_size_mb": round(file_size / 1024 / 1024, 2)
                        }
                    )

                f.write(chunk)

        logger.info(f"✅ File saved: {file_path} ({file_size / 1024:.2f} KB)")

        # 解析Excel数据
        try:
            input_data = extract_excel_data(str(file_path))
        except ValueError as e:
            # 数据验证失败，删除文件
            file_path.unlink()
            raise HTTPException(
                status_code=400,
                detail={
                    "error": "DATA_VALIDATION_FAILED",
                    "message": str(e),
                    "suggestion": "请检查Excel文件是否按照模板格式填写"
                }
            )
        except Exception as e:
            # 解析失败，删除文件
            file_path.unlink()
            logger.exception("Excel parsing failed")
            raise HTTPException(
                status_code=400,
                detail={
                    "error": "PARSE_FAILED",
                    "message": f"Excel解析失败: {str(e)}"
                }
            )

        # 保存JSON数据
        json_path = session_dir / "input.json"
        with json_path.open("w", encoding="utf-8") as f:
            json.dump(input_data, f, ensure_ascii=False, indent=2)

        logger.info(f"✅ JSON saved: {json_path}")

        # 返回成功响应
        return JSONResponse({
            "status": "success",
            "session_id": session_id,
            "filename": file.filename,
            "file_size_mb": round(file_size / 1024 / 1024, 2),
            "files": {
                "excel": str(file_path.name),
                "json": str(json_path.name)
            },
            "preview": {
                "customer_name": input_data.get("customer_name"),
                "period": input_data.get("period"),
                "alerts": {
                    "total": input_data.get("alerts", {}).get("total"),
                    "high": input_data.get("alerts", {}).get("by_severity", {}).get("high")
                },
                "categories_count": len(input_data.get("top_alert_categories", [])),
                "incidents_count": len(input_data.get("top_incidents", []))
            },
            "next_steps": {
                "description": "使用此session_id调用 /generate 端点生成报告",
                "example": {
                    "method": "POST",
                    "endpoint": "/generate",
                    "body": {
                        "input_id": "custom",  # 使用session中的input.json
                        "template_id": "mss_executive_v2",
                        "session_id": session_id
                    }
                }
            }
        })

    except HTTPException:
        # 重新抛出HTTP异常
        raise
    except Exception as e:
        # 其他未预期的错误
        logger.exception("Unexpected error during Excel upload")

        # 清理失败的文件
        if file_path.exists():
            try:
                file_path.unlink()
            except:
                pass

        raise HTTPException(
            status_code=500,
            detail={
                "error": "INTERNAL_ERROR",
                "message": f"服务器内部错误: {str(e)}"
            }
        )


# ==================== 集成示例 ====================

"""
集成到 app.py 的步骤：

1. 导入必要的依赖：
   from fastapi import UploadFile, File
   from pathlib import Path
   import json

2. 复制 extract_excel_data() 和 upload_excel() 函数

3. 在 app.py 中添加端点：

   @app.post("/upload-excel")
   async def upload_excel_endpoint(file: UploadFile = File(...)):
       return await upload_excel(
           file=file,
           session_manager=service.session_manager,
           sessions_dir=config.SESSIONS_DIR
       )

4. 测试命令：

   # 上传Excel文件
   curl -X POST http://localhost:8000/upload-excel \
     -F "file=@security_report_2025-12.xlsx"

   # 使用返回的session_id生成报告
   curl -X POST http://localhost:8000/generate \
     -H "Content-Type: application/json" \
     -d '{
       "input_id": "custom",
       "template_id": "mss_executive_v2",
       "session_id": "a3f2c5d8_20250129..."
     }'

5. 前端集成示例（HTML+JavaScript）：

   <input type="file" id="excelFile" accept=".xlsx">
   <button onclick="uploadExcel()">上传并生成报告</button>

   <script>
   async function uploadExcel() {
       const fileInput = document.getElementById('excelFile');
       const file = fileInput.files[0];

       if (!file) {
           alert('请选择文件');
           return;
       }

       const formData = new FormData();
       formData.append('file', file);

       // 上传Excel
       const uploadResp = await fetch('/upload-excel', {
           method: 'POST',
           body: formData
       });

       const uploadData = await uploadResp.json();

       if (uploadData.status === 'success') {
           console.log('上传成功:', uploadData);

           // 自动调用生成报告
           const generateResp = await fetch('/generate', {
               method: 'POST',
               headers: {'Content-Type': 'application/json'},
               body: JSON.stringify({
                   input_id: 'custom',
                   template_id: 'mss_executive_v2',
                   session_id: uploadData.session_id
               })
           });

           const generateData = await generateResp.json();
           console.log('报告生成成功:', generateData);

           // 下载报告
           window.location.href = `/download?job_id=${generateData.job_id}`;
       } else {
           alert('上传失败: ' + uploadData.message);
       }
   }
   </script>
"""
