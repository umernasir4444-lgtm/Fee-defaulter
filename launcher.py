import os
import sys
from streamlit.web import cli as stcli

def main():
    abs_app_py = os.path.abspath(os.path.join(os.path.dirname(__file__), "app.py"))
    
    # Configure streamlit arguments
    args = [
        "run",
        abs_app_py,
        "--server.port", "8765",
        "--server.address", "0.0.0.0",
        "--server.runOnSave", "true",
        "--server.fileWatcherType", "auto",
        "--browser.gatherUsageStats", "false",
        "--server.maxUploadSize", "200"
    ]
    
    # Enable HTTPS if certificates exist
    cert_path = os.path.join(os.path.dirname(__file__), "cert.pem")
    key_path = os.path.join(os.path.dirname(__file__), "key.pem")
    
    if os.path.exists(cert_path) and os.path.exists(key_path):
        args.extend([
            "--server.sslCertFile", cert_path,
            "--server.sslKeyFile", key_path
        ])
    
    sys.argv = ["streamlit"] + args
    sys.exit(stcli.main())

if __name__ == "__main__":
    main()
