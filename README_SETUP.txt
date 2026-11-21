FRESMART MULTI STORE PORTAL - FIXED LOGIN PACKAGE

LOCAL WINDOWS:
1. Install Python 3.11 or 3.12 and tick Add Python to PATH.
2. Extract the ZIP completely.
3. Double-click start_portal.bat.
4. Open http://localhost:8080/login
5. Developer login remains controlled by DEV_PWD environment variable (default from app).
6. Open Multi Store & Bridge, create store with username and temporary password.
7. Store logs in with that username/password and is forced to change password.
8. Developer can reset any created store password from Multi Store & Bridge.

IMPORTANT DATA:
- Created-store credentials are stored securely in data/portal_users.json (hashed passwords).
- Stores are stored in data/stores.json.
- Keep the data folder persistent during updates/deployment.
- On Render, attach a Persistent Disk and set DATA_DIR to its mount path (example /var/data).
- Set SECRET_KEY to a permanent random value. Do not change it on every restart.

TESTED FLOW:
Create store -> temporary-password login -> forced password change -> logout -> new-password login.
