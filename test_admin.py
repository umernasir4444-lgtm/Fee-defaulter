import sys
import os
import json
from pathlib import Path

# Add current directory to path so we can import app
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

from app import load_users, save_users, hash_password, verify_password, analyze_files, build_sample_workbook_bytes

def test_user_management():
    print("[1] Testing User Management...")
    # Setup
    users = load_users()
    test_user = "test_bot"
    test_pass = "bot_pass_123"
    
    # Add user
    users[test_user] = {"hash": hash_password(test_pass), "role": "user"}
    save_users(users)
    
    # Verify
    reloaded = load_users()
    assert test_user in reloaded, "User was not saved"
    assert verify_password(test_pass, reloaded[test_user]["hash"]), "Password verification failed"
    assert reloaded[test_user]["role"] == "user", "Role was not saved correctly"
    
    # Delete user
    del reloaded[test_user]
    save_users(reloaded)
    
    # Verify deletion
    final = load_users()
    assert test_user not in final, "User was not deleted"
    print("-> OK")

def test_analysis():
    print("[2] Testing Analysis Logic...")
    # Generate sample data
    sample_xlsx = build_sample_workbook_bytes()
    files_data = [("sample.xlsx", sample_xlsx)]
    
    # Run analysis
    analysis = analyze_files(files_data, month_name="Test Month")
    
    # Verify results
    assert analysis["student_count"] == 3, f"Expected 3 students, got {analysis['student_count']}"
    assert len(analysis["defaulters"]) == 2, f"Expected 2 defaulters, got {len(analysis['defaulters'])}"
    assert analysis["total_pending"] == 25000, f"Expected 25000 pending, got {analysis['total_pending']}"
    
    # Check individual record
    student_names = [r["student"] for r in analysis["records"]]
    assert "Ali Khan" in student_names, "Student 'Ali Khan' missing"
    print("-> OK")

if __name__ == "__main__":
    try:
        test_user_management()
        test_analysis()
        print("\n[ALL CORE TESTS PASSED SUCCESSFULLY!]")
    except Exception as e:
        print(f"[FAIL] Test encountered error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
