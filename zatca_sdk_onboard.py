"""
Onboard using SDK-generated CSR4 + verified matching key.
Key self-test confirmed VALID.

Run with:
    python manage.py shell -c "exec(open('zatca_sdk_onboard.py', encoding='utf-8').read())"

Set OTP before running.
"""
import base64
from datetime import datetime, timezone
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.backends import default_backend

# ── CONFIG ────────────────────────────────────────────────
OTP         = "702751"
DEVICE_NAME = "EGS-Main-01"
ENVIRONMENT = "simulation"
# ─────────────────────────────────────────────────────────

from invoices.models import Company, ZATCADevice
from invoices.zatca_service import ZATCAOnboarding, _der_b64_to_pem

company = Company.objects.first()
print(f"Company: {company.name}, VAT: {company.vat_number}")

# Pre-verified PKCS8 key (self-test VALID, curve=secp256k1)
KEY_PKCS8 = """-----BEGIN PRIVATE KEY-----
MIGEAgEAMBAGByqGSM49AgEGBSuBBAAKBG0wawIBAQQgzC6btySZEDa/Ula6QdPc
VlfR3DSAAQ1T1VPpzFfq2HyhRANCAASeidX5kiHMy9GWfXrjyOAh1B8+VjbNhFOD
N0XpAcFmrasTnhaYR5/7m8zsudGQD4Z06fkZaLHc+3V6Adcykomu
-----END PRIVATE KEY-----"""

# SDK-generated CSR (matches the key above)
CSR_B64 = "LS0tLS1CRUdJTiBDRVJUSUZJQ0FURSBSRVFVRVNULS0tLS0NCk1JSUI0akNDQVpjQ0FRQXdhVEVMTUFrR0ExVUVCaE1DVTBFeEN6QUpCZ05WQkFzTUFrbFVNU1V3SXdZRFZRUUsNCkRCeEJiSGRsYzNOaGJTQkRiMjUwY21GamRHbHVaeUJEYjIxd1lXNTVNU1l3SkFZRFZRUUREQjFVVTFRdE16RXgNCk5UY3hORGc0TFRNeE1UVTNNVFE0T0Rjd01EQXdNekJXTUJBR0J5cUdTTTQ5QWdFR0JTdUJCQUFLQTBJQUJKNkoNCjFmbVNJY3pMMFpaOWV1UEk0Q0hVSHo1V05zMkVVNE0zUmVrQndXYXRxeE9lRnBoSG4vdWJ6T3k1MFpBUGhuVHANCitSbG9zZHo3ZFhvQjF6S1NpYTZnZ2M0d2djc0dDU3FHU0liM0RRRUpEakdCdlRDQnVqQWhCZ2tyQmdFRUFZSTMNCkZBSUVGQXdTV2tGVVEwRXRRMjlrWlMxVGFXZHVhVzVuTUlHVUJnTlZIUkVFZ1l3d2dZbWtnWVl3Z1lNeEtUQW4NCkJnTlZCQVFNSURFdFFXeDNaWE56WVcxOE1pMHhMakI4TXkxQmJIZGxjM05oYlVWSFV6QTBNUjh3SFFZS0NaSW0NCmlaUHlMR1FCQVF3UE16RXhOVGN4TkRnNE56QXdNREF6TVEwd0N3WURWUVFNREFReE1UQXdNUTh3RFFZRFZRUWENCkRBWktaV1JrWVdneEZUQVRCZ05WQkE4TURFTnZibk4wY25WamRHbHZiakFLQmdncWhrak9QUVFEQWdOSkFEQkcNCkFpRUE5K2VCczhYUFVnVFc4a3psWVVJdk5rMWlkRHlXUlFJY216WlRqN0dmUHZBQ0lRRDA0ZDBtQnVGZE1rb2MNCjF2dnZiNWpETU5xcG90TUk5RjRvd0wzQXZ1ejJGZz09DQotLS0tLUVORCBDRVJUSUZJQ0FURSBSRVFVRVNULS0tLS0NCg=="

csr_pem = base64.b64decode(CSR_B64).decode('utf-8')
print(f"CSR loaded: {len(csr_pem)} chars")
print(f"Key loaded: {len(KEY_PKCS8)} chars")

# Verify key loads
priv = serialization.load_pem_private_key(KEY_PKCS8.encode(), password=None, backend=default_backend())
print(f"Key curve: {priv.curve.name}")

# Get compliance CSID
ob = ZATCAOnboarding(ENVIRONMENT)
compliance = ob.get_compliance_csid(csr_pem, OTP)
print(f"Compliance CSID obtained! requestID={compliance.get('requestID')}")

# Deactivate old devices
ZATCADevice.objects.filter(company=company, environment=ENVIRONMENT).update(is_active=False)

# Save device
device = ZATCADevice.objects.create(
    company=company,
    device_name=DEVICE_NAME,
    serial_number=f"{company.cr_number}-{DEVICE_NAME}",
    private_key=KEY_PKCS8,
    csr=csr_pem,
    certificate=_der_b64_to_pem(compliance['csid']),
    csid=compliance['csid'],
    secret=compliance['secret'],
    environment=ENVIRONMENT,
    is_active=True,
    registered_at=datetime.now(timezone.utc),
)
print(f"\nDevice saved: id={device.id}")
print(f"Store requestID: {compliance['requestID']}")

# Verify key matches cert
from cryptography import x509 as cx509
cert = cx509.load_pem_x509_certificate(device.certificate.encode(), default_backend())
cert_pub = cert.public_key().public_bytes(serialization.Encoding.X962, serialization.PublicFormat.UncompressedPoint).hex()
key_pub  = priv.public_key().public_bytes(serialization.Encoding.X962, serialization.PublicFormat.UncompressedPoint).hex()
print(f"Key matches cert: {cert_pub == key_pub}")

# Verify signature works
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives import hashes
sig = priv.sign(b"zatca test", ec.ECDSA(hashes.SHA256()))
try:
    cert.public_key().verify(sig, b"zatca test", ec.ECDSA(hashes.SHA256()))
    print("Signature test with cert pubkey: VALID")
except Exception as e:
    print(f"Signature test FAILED: {e}")