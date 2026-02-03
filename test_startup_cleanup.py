"""Test startup cleanup functionality with stale locks."""

import os
import time
from pathlib import Path

# Create test stale locks in outputs directory
outputs_dir = Path("mss_ai_ppt_sample_assets/backend/outputs")

print("=" * 80)
print("STARTUP CLEANUP TEST - Simulating crashed process with stale locks")
print("=" * 80)

# Create stale locks in different subdirectories
test_locks = [
    outputs_dir / "sessions" / "test_session.pptx.lock",
    outputs_dir / "reports" / "test_report.pptx.lock",
    outputs_dir / "slidespecs" / "test_spec.json.lock",
]

print("\n[1] Creating simulated stale locks (10 minutes old)...")
for lock_file in test_locks:
    lock_file.parent.mkdir(parents=True, exist_ok=True)
    lock_file.write_text(f"PID: {os.getpid()}\nCreated by test script\n")
    # Set modification time to 10 minutes ago
    old_time = time.time() - (10 * 60)
    os.utime(lock_file, (old_time, old_time))
    print(f"  [+] Created: {lock_file}")

# Count locks before cleanup
locks_before = list(outputs_dir.rglob("*.lock"))
print(f"\n[2] Locks before cleanup: {len(locks_before)}")

# Simulate startup cleanup
print("\n[3] Running startup cleanup (simulating server startup)...")
from mss_ai_ppt_sample_assets.backend.modules.file_lock import cleanup_stale_locks
cleanup_stale_locks(outputs_dir, max_age_seconds=300)  # 5 minutes

# Count locks after cleanup
locks_after = list(outputs_dir.rglob("*.lock"))
print(f"\n[4] Locks after cleanup: {len(locks_after)}")

# Verify cleanup
if len(locks_after) == 0:
    print("\n" + "=" * 80)
    print("PASS: All stale locks were cleaned up successfully")
    print("=" * 80)
    print("\nThis simulates what happens when the server starts:")
    print("  1. Process crashes while holding file locks")
    print("  2. Lock files (*.lock) remain on disk")
    print("  3. Server restarts and automatically cleans locks older than 5 minutes")
    print("  4. New requests can proceed without timeout errors")
else:
    print("\n" + "=" * 80)
    print("FAIL: Some locks were not cleaned up")
    print(f"Remaining locks: {locks_after}")
    print("=" * 80)

print("\n" + "=" * 80)
