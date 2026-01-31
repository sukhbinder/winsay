import pytest
import subprocess
import sys

@pytest.mark.skipif(sys.platform != "win32", reason="Test only runs on Windows")
def test_subprocess_call():
    """Test that say command can be called via subprocess without errors."""
    # Test calling say with subprocess - should not fail and return successful run
    try:
        result = subprocess.run([
            sys.executable, "-m", "winsay.winsay",
            "testing in progress", "-v", "10"
        ], capture_output=True, text=True, timeout=10)

        # The command should execute successfully (exit code 0)
        assert result.returncode == 0
    except subprocess.TimeoutExpired:
        pytest.fail("Subprocess call timed out")
    except Exception as e:
        pytest.fail(f"Subprocess call failed with error: {e}")
