"""Add progress notifications to generate endpoint"""

import re

app_file = "mss_ai_ppt_sample_assets/backend/app.py"

with open(app_file, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Remove duplicate registration (fix bug)
content = re.sub(
    r'    # Register WebSocket connection if client_id provided\n    if req\.client_id and req\.session_id:\n        ws_manager\.register_session\(req\.session_id, req\.client_id\)\n\n    # Register WebSocket connection if client_id provided\n    if req\.client_id and req\.session_id:\n        ws_manager\.register_session\(req\.session_id, req\.client_id\)',
    '''    # Register WebSocket connection if client_id provided
    if req.client_id and req.session_id:
        ws_manager.register_session(req.session_id, req.client_id)
        # Send initial progress
        await ws_manager.send_progress_update(
            req.session_id, 0, "请求已接收,正在排队...",
            {"template": req.template_id}
        )''',
    content
)

# 2. Add progress after acquiring slot
content = content.replace(
    '        logger.info(f"✅ Acquired LLM slot, starting generation...")',
    '''        logger.info(f"✅ Acquired LLM slot, starting generation...")

        # Send progress: started generation
        if req.client_id and req.session_id:
            await ws_manager.send_progress_update(
                req.session_id, 10, "已获得处理槽位,开始生成报告..."
            )'''
)

# 3. Add progress before execution
content = content.replace(
    '            # Run the synchronous generate() in a thread pool to avoid blocking\n            loop = asyncio.get_running_loop()',
    '''            # Send progress: executing
            if req.client_id and req.session_id:
                await ws_manager.send_progress_update(
                    req.session_id, 30, "正在调用AI生成内容,请稍候..."
                )

            # Run the synchronous generate() in a thread pool to avoid blocking
            loop = asyncio.get_running_loop()'''
)

# 4. Add completion notification
content = content.replace(
    '            logger.info(f"✓ Generation successful: {result.get(\'job_id\')}")',
    '''            logger.info(f"✓ Generation successful: {result.get('job_id')}")

            # Send completion notification
            if req.client_id and req.session_id:
                await ws_manager.send_completion(
                    req.session_id,
                    result,
                    success=True
                )'''
)

# 5. Add failure notification
content = content.replace(
    '        except Exception as e:\n            logger.exception(f"✗ Generation failed with exception: {e}")\n            raise HTTPException(status_code=500, detail=str(e))',
    '''        except Exception as e:
            logger.exception(f"✗ Generation failed with exception: {e}")

            # Send failure notification
            if req.client_id and req.session_id:
                await ws_manager.send_completion(
                    req.session_id,
                    {"error": str(e)},
                    success=False
                )

            raise HTTPException(status_code=500, detail=str(e))'''
)

with open(app_file, 'w', encoding='utf-8') as f:
    f.write(content)

print("Progress notifications added:")
print("  [0%]  Request received")
print("  [10%] Acquired processing slot")
print("  [30%] Calling AI generation")
print("  [100%] Completion/failure")
print("")
print("Restart server to test!")
