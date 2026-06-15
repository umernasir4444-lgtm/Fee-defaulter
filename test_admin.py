import urllib.request
import urllib.parse
import http.cookiejar
import threading
import time
import sys
import json
from app import ThreadingHTTPServer, Handler, load_users, save_users, hash_password

PORT = 9999

def run_server():
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    server.serve_forever()

def main():
    # Force clean users.json for testing
    users = {"umer": {"hash": hash_password("123"), "role": "admin"}}
    save_users(users)

    # Start server in background thread
    t = threading.Thread(target=run_server, daemon=True)
    t.start()
    time.sleep(1) # wait for startup

    # Setup cookie jar for session tracking
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    urllib.request.install_opener(opener)

    print("[1] Test Login GET...")
    res = urllib.request.urlopen(f"http://127.0.0.1:{PORT}/login")
    html_content = res.read().decode('utf-8')
    assert "Login" in html_content, "Login form not found"
    print("-> OK")

    print("[2] Test Login POST...")
    data = urllib.parse.urlencode({"username": "umer", "password": "123"}).encode()
    res = urllib.request.urlopen(f"http://127.0.0.1:{PORT}/login", data=data)
    html_content = res.read().decode('utf-8')
    assert "Fee Defaulter Report Generator" in html_content, "Dashboard not loaded after login"
    # Verify Admin Panel button is in dashboard header
    assert "Admin Panel" in html_content, "Admin Panel button missing for umer"
    print("-> OK")

    print("[3] Test GET /admin...")
    res = urllib.request.urlopen(f"http://127.0.0.1:{PORT}/admin")
    html_content = res.read().decode('utf-8')
    assert "Admin Panel" in html_content, "Admin panel page not loaded"
    assert "Role" in html_content, "Role column/label missing"
    print("-> OK")

    print("[4] Test POST /admin (Add User with role)...")
    data = urllib.parse.urlencode({"username": "testuser", "password": "abc", "role": "user", "action": "add"}).encode()
    res = urllib.request.urlopen(f"http://127.0.0.1:{PORT}/admin", data=data)
    html_content = res.read().decode('utf-8')
    assert "added/updated successfully" in html_content, "User addition feedback missing"
    assert "Delete User" in html_content, "Delete User button missing after adding user"
    
    # Check users.json to ensure testuser is added and has role 'user'
    current_users = load_users()
    assert "testuser" in current_users, "testuser not found in database"
    assert current_users["testuser"]["role"] == "user", "testuser has wrong role"
    print("-> OK")

    print("[5] Test GET /impersonate (Switch Dashboard)...")
    res = urllib.request.urlopen(f"http://127.0.0.1:{PORT}/impersonate?username=testuser")
    html_content = res.read().decode('utf-8')
    assert "Fee Defaulter Report Generator" in html_content, "Dashboard not loaded after impersonation"
    # Ensure Admin Panel button is NOT visible for standard 'user'
    assert "Admin Panel" not in html_content, "Admin Panel button visible for standard user"
    print("-> OK")

    print("[6] Switch back to umer...")
    data = urllib.parse.urlencode({"username": "umer", "password": "123"}).encode()
    urllib.request.urlopen(f"http://127.0.0.1:{PORT}/login", data=data)
    print("-> OK")

    print("[7] Test POST /admin (Delete User)...")
    data = urllib.parse.urlencode({"username": "testuser", "action": "remove"}).encode()
    res = urllib.request.urlopen(f"http://127.0.0.1:{PORT}/admin", data=data)
    html_content = res.read().decode('utf-8')
    assert "deleted successfully" in html_content, "User deletion feedback missing"
    
    # Check users.json to ensure testuser is deleted
    current_users = load_users()
    assert "testuser" not in current_users, "testuser still exists in database"
    print("-> OK")

    print("\n[ALL TESTS PASSED SUCCESSFULLY!]")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[FAIL] Test encountered error: {e}")
        sys.exit(1)
