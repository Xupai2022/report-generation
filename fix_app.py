"""Quick fix script for app.py bugs.

Fixes:
1. AttributeError on line 133: _waiters.__len__() when _waiters is None
2. asyncio.get_event_loop() -> asyncio.get_running_loop() on line 139
"""

import re

APP_PY_PATH = "mss_ai_ppt_sample_assets/backend/app.py"

# Read file
with open(APP_PY_PATH, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Replace the buggy logging line
old_log = 'logger.info(f"🔄 Waiting for LLM slot (current waiting: {llm_semaphore._waiters.__len__() if hasattr(llm_semaphore, \'_waiters\') else 0})")'
new_log = '''waiting_count = 0
    if hasattr(llm_semaphore, '_waiters') and llm_semaphore._waiters is not None:
        waiting_count = len(llm_semaphore._waiters)
    logger.info(f"🔄 Waiting for LLM slot (current waiting: {waiting_count})")'''

content = content.replace(old_log, new_log)

# Fix 2: Replace get_event_loop with get_running_loop
content = content.replace(
    'loop = asyncio.get_event_loop()',
    'loop = asyncio.get_running_loop()'
)

# Write back
with open(APP_PY_PATH, 'w', encoding='utf-8') as f:
    f.write(content)

print("✅ Fixed app.py:")
print("  1. Fixed _waiters.__len__() AttributeError")
print("  2. Changed get_event_loop() -> get_running_loop()")
print("\n🔄 Please restart the server:")
print("  1. Kill current process")
print("  2. Run: python -m mss_ai_ppt_sample_assets.backend.app")
