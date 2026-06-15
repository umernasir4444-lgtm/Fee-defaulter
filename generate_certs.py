import os
from pathlib import Path

def generate_self_signed_cert(cert_path, key_path):
    try:
        from cryptography import x509
        from cryptography.x509.oid import NameOID
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.hazmat.primitives import serialization
        import datetime

        print("[INFO] Generating self-signed certificate using 'cryptography'...")
        
        key = rsa.generate_private_key(
            public_exponent=65537,
            key_size=2048,
        )
        
        subject = issuer = x509.Name([
            x509.NameAttribute(NameOID.COUNTRY_NAME, u"US"),
            x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, u"CA"),
            x509.NameAttribute(NameOID.LOCALITY_NAME, u"San Francisco"),
            x509.NameAttribute(NameOID.ORGANIZATION_NAME, u"Fee Generator"),
            x509.NameAttribute(NameOID.COMMON_NAME, u"localhost"),
        ])
        
        cert = x509.CertificateBuilder().subject_name(
            subject
        ).issuer_name(
            issuer
        ).public_key(
            key.public_key()
        ).serial_number(
            x509.random_serial_number()
        ).not_valid_before(
            datetime.datetime.utcnow()
        ).not_valid_after(
            datetime.datetime.utcnow() + datetime.timedelta(days=3650)
        ).add_extension(
            x509.SubjectAlternativeName([x509.DNSName(u"localhost")]),
            critical=False,
        ).sign(key, hashes.SHA256())

        with open(key_path, "wb") as f:
            f.write(key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.TraditionalOpenSSL,
                encryption_algorithm=serialization.NoEncryption(),
            ))
            
        with open(cert_path, "wb") as f:
            f.write(cert.public_bytes(serialization.Encoding.PEM))
            
        print(f"[OK] Certificate and key generated: {cert_path}, {key_path}")
        return True
    except ImportError:
        print("[INFO] 'cryptography' library not found. Falling back to simple dummy cert (if openssl is available) or skipping HTTPS.")
        # Attempting a shell call to openssl as a fallback
        import subprocess
        try:
            cmd = f'openssl req -x509 -newkey rsa:2048 -keyout "{key_path}" -out "{cert_path}" -days 3650 -nodes -subj "/CN=localhost"'
            subprocess.run(cmd, shell=True, check=True, capture_output=True)
            print(f"[OK] Certificate generated via OpenSSL.")
            return True
        except:
            print("[WARN] Could not generate certificates. HTTPS will be disabled.")
            return False

if __name__ == "__main__":
    APP_DIR = Path(__file__).resolve().parent
    generate_self_signed_cert(APP_DIR / "cert.pem", APP_DIR / "key.pem")
