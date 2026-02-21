"""Test script for validating improvements to the codebase.

Tests:
1. Exception hierarchy
2. Session ID validation (new format only)
3. Excel handler
4. Module imports
"""

import sys
from pathlib import Path

# Add backend to path
backend_path = Path(__file__).parent / "mss_ai_ppt_sample_assets" / "backend"
sys.path.insert(0, str(backend_path.parent))

def test_exceptions():
    """Test the exception hierarchy."""
    print("\n=== Testing Exception Hierarchy ===")

    from mss_ai_ppt_sample_assets.backend.exceptions import (
        MSSAIException,
        InputNotFoundError,
        TemplateNotFoundError,
        InvalidSessionIDError,
        FileValidationError,
        DataValidationError,
        LLMGenerationError,
        FileLockError,
    )

    # Test base exception
    try:
        raise MSSAIException("Test error", error_code="TEST_ERROR", context={"key": "value"})
    except MSSAIException as e:
        assert e.message == "Test error"
        assert e.error_code == "TEST_ERROR"
        assert e.context["key"] == "value"
        assert "error" in e.to_dict()
        print("✅ Base exception works correctly")

    # Test specific exceptions
    try:
        raise InputNotFoundError("test_input")
    except InputNotFoundError as e:
        assert "test_input" in str(e)
        assert e.error_code == "INPUT_NOT_FOUND"
        print("✅ InputNotFoundError works correctly")

    try:
        raise FileValidationError("test.xlsx", "invalid format")
    except FileValidationError as e:
        assert "test.xlsx" in str(e)
        print("✅ FileValidationError works correctly")

    print("✅ All exception tests passed!")


def test_session_validation():
    """Test session ID validation with new format only."""
    print("\n=== Testing Session ID Validation ===")

    from mss_ai_ppt_sample_assets.backend.modules.session_manager import SessionManager
    import tempfile

    # Create temporary session manager
    with tempfile.TemporaryDirectory() as tmpdir:
        sm = SessionManager(Path(tmpdir))

        # Test valid new format
        valid_ids = [
            "a3f2c5d8_20250129143025123456",
            "1234abcd_20260130171530999999",
            "ffff0000_19990101000000000000",
        ]

        for session_id in valid_ids:
            try:
                sm.validate_session_id(session_id)
                print(f"✅ Valid ID accepted: {session_id}")
            except ValueError as e:
                print(f"❌ Valid ID rejected: {session_id} - {e}")
                raise

        # Test invalid IDs
        invalid_ids = [
            "session_1769764397059_1qz5m42dv",  # Old format (should be rejected)
            "../etc/passwd",  # Path traversal
            "abc/123_20250129143025123456",  # Contains slash
            "abc\\123_20250129143025123456",  # Contains backslash
            "abc..123_20250129143025123456",  # Contains ..
            "toolong12_20250129143025123456",  # UUID part too long
            "short_20250129143025123456",  # UUID part too short
            "ABCD1234_20250129143025123456",  # Uppercase (should be lowercase)
            "a3f2c5d8_2025012914302512345",  # Timestamp too short
            "",  # Empty
        ]

        for session_id in invalid_ids:
            try:
                sm.validate_session_id(session_id)
                print(f"❌ Invalid ID accepted: {session_id}")
                raise AssertionError(f"Should have rejected: {session_id}")
            except ValueError:
                print(f"✅ Invalid ID rejected: {session_id}")

        # Test session ID generation
        generated_id = sm.generate_session_id()
        print(f"✅ Generated session ID: {generated_id}")
        sm.validate_session_id(generated_id)
        print("✅ Generated ID passes validation")

    print("✅ All session validation tests passed!")


def test_excel_handler():
    """Test Excel handler module."""
    print("\n=== Testing Excel Handler ===")

    from mss_ai_ppt_sample_assets.backend.modules.excel_handler import (
        ExcelValidator,
        ExcelDataExtractor,
        ExcelHandler,
    )
    from mss_ai_ppt_sample_assets.backend.exceptions import FileValidationError

    # Test validator
    validator = ExcelValidator(max_size_mb=10)

    # Valid extension
    try:
        validator.validate_extension("test.xlsx")
        print("✅ Valid .xlsx extension accepted")
    except FileValidationError:
        raise AssertionError("Should accept .xlsx")

    # Invalid extensions
    for ext in [".xlsm", ".xls", ".csv", ".txt"]:
        try:
            validator.validate_extension(f"test{ext}")
            raise AssertionError(f"Should reject {ext}")
        except FileValidationError:
            print(f"✅ Invalid extension {ext} rejected")

    # Test MIME type validation
    try:
        validator.validate_mime_type(
            "test.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        print("✅ Valid MIME type accepted")
    except FileValidationError:
        raise AssertionError("Should accept valid MIME type")

    try:
        validator.validate_mime_type("test.xlsx", "text/plain")
        raise AssertionError("Should reject invalid MIME type")
    except FileValidationError:
        print("✅ Invalid MIME type rejected")

    # Test size validation
    try:
        validator.validate_size("test.xlsx", 5 * 1024 * 1024)  # 5MB
        print("✅ Valid file size accepted")
    except FileValidationError:
        raise AssertionError("Should accept valid size")

    try:
        validator.validate_size("test.xlsx", 15 * 1024 * 1024)  # 15MB
        raise AssertionError("Should reject oversized file")
    except FileValidationError:
        print("✅ Oversized file rejected")

    print("✅ All Excel handler tests passed!")


def test_module_imports():
    """Test that all new modules can be imported."""
    print("\n=== Testing Module Imports ===")

    try:
        from mss_ai_ppt_sample_assets.backend import exceptions
        print("✅ exceptions module imported")

        from mss_ai_ppt_sample_assets.backend.modules import excel_handler
        print("✅ excel_handler module imported")

        from mss_ai_ppt_sample_assets.backend.modules import session_manager
        print("✅ session_manager module imported")

        from mss_ai_ppt_sample_assets.backend import websocket_support
        print("✅ websocket_support module imported")

        print("✅ All module imports successful!")

    except ImportError as e:
        print(f"❌ Import failed: {e}")
        raise


def main():
    """Run all tests."""
    print("=" * 60)
    print("Testing Codebase Improvements")
    print("=" * 60)

    try:
        test_module_imports()
        test_exceptions()
        test_session_validation()
        test_excel_handler()

        print("\n" + "=" * 60)
        print("✅ ALL TESTS PASSED!")
        print("=" * 60)
        return 0

    except Exception as e:
        print("\n" + "=" * 60)
        print(f"❌ TESTS FAILED: {e}")
        print("=" * 60)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
