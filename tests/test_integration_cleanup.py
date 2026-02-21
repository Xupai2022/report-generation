"""Quick integration test for startup cleanup functionality."""

import time
import os
from pathlib import Path

print("=" * 80)
print("INTEGRATION TEST: Startup Cleanup with Stale Locks")
print("=" * 80)

# Step 1: Create stale locks
outputs_dir = Path("mss_ai_ppt_sample_assets/backend/outputs")
test_locks = [
    outputs_dir / "sessions" / "integration_test.lock",
    outputs_dir / "reports" / "integration_test.lock",
]

print("\n[Step 1] Creating stale locks...")
for lock_file in test_locks:
    lock_file.parent.mkdir(parents=True, exist_ok=True)
    lock_file.write_text(f"PID: {os.getpid()}\n")
    # Set to 10 minutes old
    old_time = time.time() - (10 * 60)
    os.utime(lock_file, (old_time, old_time))
    print(f"  Created: {lock_file}")

# Step 2: Import and run cleanup (simulating startup)
print("\n[Step 2] Simulating server startup cleanup...")
from mss_ai_ppt_sample_assets.backend.modules.file_lock import cleanup_stale_locks
cleanup_stale_locks(outputs_dir, max_age_seconds=300)

# Step 3: Verify cleanup
print("\n[Step 3] Verifying cleanup...")
remaining = [lock for lock in test_locks if lock.exists()]

if len(remaining) == 0:
    print("\n" + "=" * 80)
    print("SUCCESS: Integration test passed")
    print("=" * 80)
    print("\nStartup cleanup functionality is working correctly:")
    print("  - Stale locks are automatically detected")
    print("  - Locks older than 5 minutes are removed")
    print("  - System can recover from process crashes")
else:
    print("\n" + "=" * 80)
    print("FAILURE: Integration test failed")
    print("=" * 80)
    print(f"Remaining locks: {remaining}")

print("\n" + "=" * 80)
