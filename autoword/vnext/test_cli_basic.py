#!/usr/bin/env python3
"""
Basic CLI functionality test for AutoWord vNext.

This script tests the CLI interface without requiring actual document processing.
"""

import sys
import subprocess
import tempfile
import os
from pathlib import Path


def run_cli_command(args):
    """Run CLI command and return result."""
    cmd = [sys.executable, "-m", "autoword.vnext.cli"] + args
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        return result.returncode, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return -1, "", "Command timed out"
    except Exception as e:
        return -1, "", str(e)


def test_help():
    """Test help command."""
    print("Testing help command...")
    returncode, stdout, stderr = run_cli_command(["--help"])
    
    if returncode != 0:
        print(f"❌ Help command failed: {stderr}")
        return False
    
    if "AutoWord vNext Pipeline CLI" not in stdout:
        print("❌ Help text missing expected content")
        return False
    
    print("✅ Help command works")
    return True


def test_status():
    """Test status command."""
    print("Testing status command...")
    returncode, stdout, stderr = run_cli_command(["status"])
    
    if "System Status" not in stdout:
        print("❌ Status command missing expected output")
        return False
    
    print("✅ Status command works")
    return True


def test_config_show():
    """Test config show command."""
    print("Testing config show command...")
    returncode, stdout, stderr = run_cli_command(["config", "show"])
    
    if returncode != 0:
        print(f"❌ Config show failed: {stderr}")
        return False
    
    if "Current Configuration" not in stdout:
        print("❌ Config show missing expected output")
        return False
    
    print("✅ Config show command works")
    return True


def test_config_create():
    """Test config create command."""
    print("Testing config create command...")
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        temp_config = f.name
    
    try:
        returncode, stdout, stderr = run_cli_command(["config", "create", temp_config])
        
        if returncode != 0:
            print(f"❌ Config create failed: {stderr}")
            return False
        
        if not os.path.exists(temp_config):
            print("❌ Config file was not created")
            return False
        
        # Check if it's valid JSON
        import json
        with open(temp_config, 'r') as f:
            config = json.load(f)
        
        if "model" not in config or "monitoring_level" not in config:
            print("❌ Config file missing expected keys")
            return False
        
        print("✅ Config create command works")
        return True
        
    finally:
        if os.path.exists(temp_config):
            os.unlink(temp_config)


def test_invalid_command():
    """Test invalid command handling."""
    print("Testing invalid command handling...")
    returncode, stdout, stderr = run_cli_command(["invalid-command"])
    
    if returncode == 0:
        print("❌ Invalid command should return non-zero exit code")
        return False
    
    print("✅ Invalid command handling works")
    return True


def main():
    """Run all CLI tests."""
    print("=== AutoWord vNext CLI Basic Tests ===\n")
    
    tests = [
        test_help,
        test_status,
        test_config_show,
        test_config_create,
        test_invalid_command
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        try:
            if test():
                passed += 1
            print()  # Add spacing between tests
        except Exception as e:
            print(f"❌ Test {test.__name__} failed with exception: {e}\n")
    
    print(f"=== Test Results ===")
    print(f"Passed: {passed}/{total}")
    
    if passed == total:
        print("🎉 All CLI tests passed!")
        return 0
    else:
        print("❌ Some CLI tests failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())