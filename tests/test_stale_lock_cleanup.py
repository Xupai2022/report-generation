"""Test stale lock cleanup functionality."""

import os
import time
from pathlib import Path
from mss_ai_ppt_sample_assets.backend.modules.file_lock import cleanup_stale_locks

# Create test directory
test_dir = Path("test_locks")
test_dir.mkdir(exist_ok=True)

print("=" * 80)
print("STALE LOCK CLEANUP TEST")
print("=" * 80)

# Create fresh lock file (< 5 minutes old)
fresh_lock = test_dir / "fresh.lock"
fresh_lock.write_text("PID: 12345")
print(f"[+] Created fresh lock: {fresh_lock}")

# Create old lock file (simulated as > 5 minutes old)
old_lock = test_dir / "old.lock"
old_lock.write_text("PID: 67890")
# Modify its modification time to 10 minutes ago
old_time = time.time() - (10 * 60)  # 10 minutes ago
os.utime(old_lock, (old_time, old_time))
print(f"[+] Created old lock (10 min ago): {old_lock}")

# Create subdirectory with stale lock
sub_dir = test_dir / "subdir"
sub_dir.mkdir(exist_ok=True)
sub_lock = sub_dir / "sub.lock"
sub_lock.write_text("PID: 99999")
sub_time = time.time() - (20 * 60)  # 20 minutes ago
os.utime(sub_lock, (sub_time, sub_time))
print(f"[+] Created old lock in subdir (20 min ago): {sub_lock}")

print("\nBefore cleanup:")
print(f"  Total locks: {len(list(test_dir.rglob('*.lock')))}")

# Run cleanup (max_age_seconds=300 = 5 minutes)
print("\nRunning cleanup (max_age: 5 minutes)...")
cleanup_stale_locks(test_dir, max_age_seconds=300)

# Check results
remaining_locks = list(test_dir.rglob("*.lock"))
print("\nAfter cleanup:")
print(f"  Remaining locks: {len(remaining_locks)}")

# Verify results
expected_lock = fresh_lock
if len(remaining_locks) == 1 and remaining_locks[0] == expected_lock:
    print("\n" + "=" * 80)
    print("PASS: Only fresh lock remains (old locks cleaned up)")
    print("=" * 80)
else:
    print("\n" + "=" * 80)
    print("FAIL: Unexpected cleanup result")
    print(f"Expected: ['{expected_lock}']")
    print(f"Got: {remaining_locks}")
    print("=" * 80)

# Cleanup test directory
import shutil
shutil.rmtree(test_dir)
print(f"\nCleaned up test directory: {test_dir}")
