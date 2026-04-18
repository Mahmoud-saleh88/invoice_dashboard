"""
=============================================================
 ZATCA e-Invoicing Integration Service — Phase 2
 Round-7 — Production-ready with SDK signing.

 Key changes over Round-6:
   [SDK-SIGN]     ZATCASDKSigner uses the ZATCA Java SDK for signing
                  (proven to pass ZATCA validation vs Python signer).
   [ONBOARD-FIX]  onboard_device skips production CSID for non-production
                  environments (simulation only supports compliance CSID).
   [ADDR-FIX]     validate_invoice checks both seller AND buyer address.
   [CLIENT-FIX]   effective_street_name no longer references missing 'address'
                  field — uses street_name only.
   [DELIVERY-FIX] _delivery() added to XML generator (KSA-5 ActualDeliveryDate).
   [WARN-FIX]     Address warnings fixed by enforcing district/building/postal
                  in both Company and Client before submission.

=============================================================
"""

from __future__ import annotations

import base64
import hashlib
import json
import logging
import os
import re as _re
import subprocess
import uuid
from datetime import datetime, timezone
from io import BytesIO
from typing import Optional

import qrcode
import requests
from cryptography import x509 as cx509
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.x509.oid import NameOID
from lxml import etree
from tenacity import (
    retry, retry_if_exception_type,
    stop_after_attempt, wait_exponential,
)
import os
from django.conf import settings
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────────────────
# SDK Configuration — update this path for your server
# ──────────────────────────────────────────────────────────
# Windows VPS / local development:


SDK_DIR = os.path.join(settings.BASE_DIR,"zatca_sdk","zatca-einvoicing-sdk-Java-238-R3.4.8")
SDK_JAR = os.path.join(SDK_DIR, "Apps", "zatca-einvoicing-sdk-238-R3.4.8.jar")

# ──────────────────────────────────────────────────────────
# API Endpoints
# ──────────────────────────────────────────────────────────
ZATCA_SANDBOX_BASE = "https://gw-fatoora.zatca.gov.sa/e-invoicing/developer-portal"
ZATCA_SIMULATION_BASE = "https://gw-fatoora.zatca.gov.sa/e-invoicing/simulation"
ZATCA_PROD_BASE = "https://gw-fatoora.zatca.gov.sa/e-invoicing/core"

ENDPOINTS = {
    "compliance":          "/compliance",
    "production_csids":    "/production/csids",
    "compliance_invoices": "/compliance/invoices",
    "clearance":           "/invoices/clearance/single",
    "reporting":           "/invoices/reporting/single",
}

# Official ZATCA genesis hash for the very first invoice (KSA-13)
ZATCA_GENESIS_HASH = "NWZlMmI4YTNjZWJhNzgxNzNhNGI3NWI3YTMyNDY5ZGQ="

# ──────────────────────────────────────────────────────────
# Namespace map
# ──────────────────────────────────────────────────────────
NS = {
    None:    "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
    "cac":   "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cbc":   "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "ext":   "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2",
    "sig":   "urn:oasis:names:specification:ubl:schema:xsd:CommonSignatureComponents-2",
    "sac":   "urn:oasis:names:specification:ubl:schema:xsd:SignatureAggregateComponents-2",
    "sbc":   "urn:oasis:names:specification:ubl:schema:xsd:SignatureBasicComponents-2",
    "ds":    "http://www.w3.org/2000/09/xmldsig#",
    "xades": "http://uri.etsi.org/01903/v1.3.2#",
}

_EXT   = NS["ext"]
_CAC   = NS["cac"]
_CBC   = NS["cbc"]
_DS    = NS["ds"]
_XADES = NS["xades"]


# ──────────────────────────────────────────────────────────
# Exceptions
# ──────────────────────────────────────────────────────────

class ZATCAError(Exception):
    """Base ZATCA exception."""

class ZATCAValidationError(ZATCAError):
    """Invoice failed pre-submission checks."""

class ZATCACommunicationError(ZATCAError):
    """Network / API error."""


# ──────────────────────────────────────────────────────────
# Key generation & CSR
# ──────────────────────────────────────────────────────────

def generate_ec_keypair() -> tuple[str, str]:
    """Generate secp256k1 key pair in PKCS#8 format."""
    private_key = ec.generate_private_key(ec.SECP256K1(), default_backend())
    private_pem = private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption()
    )
    public_pem = private_key.public_key().public_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PublicFormat.SubjectPublicKeyInfo
    )
    return private_pem.decode(), public_pem.decode()


def generate_csr(
    company_name: str,
    vat_number: str,
    serial_number: str,
    organization_unit: str = "IT",
    country: str = "SA",
    private_key_pem: Optional[str] = None,
    invoice_type: str = "1100",
    address: str = "Riyadh",
    business_category: str = "Contracting",
) -> tuple[str, str]:
    """Generate a ZATCA-compliant CSR. Returns (csr_pem, private_key_pem)."""
    if private_key_pem:
        priv = serialization.load_pem_private_key(
            private_key_pem.encode(), password=None, backend=default_backend()
        )
    else:
        priv_pem, _ = generate_ec_keypair()
        priv = serialization.load_pem_private_key(
            priv_pem.encode(), password=None, backend=default_backend()
        )
        private_key_pem = priv_pem

    vat_number = vat_number.strip()
    if len(vat_number) != 15 or not vat_number.isdigit():
        raise ValueError(f"VAT number must be exactly 15 digits, got: '{vat_number}'")

    egs_serial = f"1-{organization_unit}|2-EGS|3-{serial_number}"
    serial_short = serial_number.split("-")[-1][:20]
    cn = f"TST-{serial_short}-{vat_number}"

    template_bytes = b"ZATCA-Code-Signing"
    utf8_der = bytes([0x0C, len(template_bytes)]) + template_bytes
    cert_template_der = bytes([0x30, len(utf8_der)]) + utf8_der

    OID_CERT_TEMPLATE = cx509.ObjectIdentifier("1.3.6.1.4.1.311.20.2")
    OID_UID           = cx509.ObjectIdentifier("0.9.2342.19200300.100.1.1")
    OID_SERIAL        = cx509.ObjectIdentifier("2.5.4.5")
    OID_TITLE         = cx509.ObjectIdentifier("2.5.4.12")
    OID_REG_ADDR      = cx509.ObjectIdentifier("2.5.4.26")
    OID_BUS_CAT       = cx509.ObjectIdentifier("2.5.4.15")

    san_dn = cx509.Name([
        cx509.NameAttribute(OID_SERIAL,   egs_serial),
        cx509.NameAttribute(OID_UID,      vat_number),
        cx509.NameAttribute(OID_TITLE,    invoice_type),
        cx509.NameAttribute(OID_REG_ADDR, address),
        cx509.NameAttribute(OID_BUS_CAT,  business_category),
    ])

    csr = (
        cx509.CertificateSigningRequestBuilder()
        .subject_name(cx509.Name([
            cx509.NameAttribute(NameOID.COUNTRY_NAME,             country),
            cx509.NameAttribute(NameOID.ORGANIZATIONAL_UNIT_NAME, organization_unit[:64]),
            cx509.NameAttribute(NameOID.ORGANIZATION_NAME,        company_name[:64]),
            cx509.NameAttribute(NameOID.COMMON_NAME,              cn[:64]),
        ]))
        .add_extension(
            cx509.UnrecognizedExtension(OID_CERT_TEMPLATE, cert_template_der),
            critical=False,
        )
        .add_extension(
            cx509.SubjectAlternativeName([cx509.DirectoryName(san_dn)]),
            critical=False,
        )
        .sign(priv, hashes.SHA256(), default_backend())
    )

    csr_pem = csr.public_bytes(serialization.Encoding.PEM).decode()
    log.debug("CSR generated -- CN=%s VAT=%s len=%d", cn, vat_number, len(csr_pem))
    return csr_pem, private_key_pem


# ──────────────────────────────────────────────────────────
# ZATCA Onboarding
# ──────────────────────────────────────────────────────────

class ZATCAOnboarding:

    def __init__(self, environment: str = "sandbox"):
        if environment == "production":
            self.base_url = ZATCA_PROD_BASE
        elif environment == "simulation":
            self.base_url = ZATCA_SIMULATION_BASE
        else:
            self.base_url = ZATCA_SANDBOX_BASE

    def _h(self, *, otp=None, username=None, password=None):
        h = {"Content-Type": "application/json",
             "Accept": "application/json",
             "Accept-Version": "V2"}
        if otp:
            h["OTP"] = otp
        if username and password:
            token = base64.b64encode(f"{username}:{password}".encode()).decode()
            h["Authorization"] = f"Basic {token}"
        return h

    def get_compliance_csid(self, csr_pem: str, otp: str) -> dict:
        otp = ''.join(c for c in otp if c.isdigit())
        if len(otp) != 6:
            raise ZATCAError(f"OTP must be 6 digits, got {len(otp)} digits")

        csr_base64 = ''.join(base64.b64encode(csr_pem.encode()).decode().split())
        payload = {"csr": csr_base64}
        headers = {
            "Accept": "application/json",
            "Accept-Language": "en",
            "Accept-Version": "V2",
            "Content-Type": "application/json",
            "OTP": otp,
        }

        url = f"{self.base_url}{ENDPOINTS['compliance']}"
        log.info("Requesting compliance CSID from %s", url)

        try:
            resp = requests.post(url, json=payload, headers=headers, timeout=30)
        except requests.exceptions.RequestException as e:
            raise ZATCACommunicationError(f"Network error: {e}")

        if not resp.ok:
            msg = self._extract_error_msg(resp)
            raise ZATCAError(
                f"Compliance CSID API HTTP {resp.status_code}: {msg}\n"
                f"Full body: {resp.text[:800]}"
            )

        d = resp.json()
        token = d.get("binarySecurityToken") or d.get("BinarySecurityToken")
        if not token:
            raise ZATCAError(
                f"No binarySecurityToken in response. "
                f"dispositionMessage: {d.get('dispositionMessage')}\n"
                f"Full response: {d}"
            )

        return {
            "requestID":          d.get("requestID"),
            "csid":               token,
            "secret":             d.get("secret"),
            "dispositionMessage": d.get("dispositionMessage"),
        }

    @staticmethod
    def _extract_error_msg(resp) -> str:
        try:
            body = resp.json()
            if body.get("dispositionMessage"):
                return body["dispositionMessage"]
            if body.get("message"):
                return body["message"]
            for key in ("errors", "errorMessages"):
                errs = body.get(key, [])
                if errs and isinstance(errs, list):
                    return " | ".join(
                        f"{e.get('code','?')}: {e.get('message', str(e))}"
                        for e in errs[:3]
                    )
            vr = body.get("validationResults", {})
            errs = vr.get("errorMessages", [])
            if errs:
                return " | ".join(
                    f"{e.get('code','?')}: {e.get('message','?')}"
                    for e in errs[:3]
                )
            return str(body)[:400]
        except Exception:
            return resp.text[:400]

    def get_production_csid(self, compliance_csid, compliance_secret,
                            compliance_request_id) -> dict:
        payload = {"compliance_request_id": compliance_request_id}
        resp = requests.post(
            f"{self.base_url}{ENDPOINTS['production_csids']}",
            json=payload,
            headers=self._h(username=compliance_csid, password=compliance_secret),
            timeout=30,
        )
        resp.raise_for_status()
        d = resp.json()
        return {
            "csid":               d.get("binarySecurityToken"),
            "secret":             d.get("secret"),
            "dispositionMessage": d.get("dispositionMessage"),
        }


# ──────────────────────────────────────────────────────────
# SDK Signer — uses ZATCA Java SDK (PROVEN TO WORK)
# ──────────────────────────────────────────────────────────

class ZATCASDKSigner:
    """
    Signs invoices using the official ZATCA Java SDK.
    This is the ONLY signing method proven to pass ZATCA validation.
    The Java SDK must be installed on the same machine as the Django app.
    """

    def __init__(self, device, sdk_dir: str = None, sdk_jar: str = None):
        self.device  = device
        self.sdk_dir = sdk_dir or SDK_DIR
        self.sdk_jar = sdk_jar or SDK_JAR

    def _write_cert_and_key(self) -> tuple[str, str]:
        """Write cert and key in SDK-expected format (raw base64, no PEM headers)."""
        # Certificate: strip PEM headers → raw base64
        cert_lines = [
            l for l in self.device.certificate.strip().split('\n')
            if not l.startswith('-----')
        ]
        cert_raw = ''.join(cert_lines)

        # Key: PKCS8 → SEC1 → strip headers → raw base64
        priv = serialization.load_pem_private_key(
            self.device.private_key.encode(), password=None, backend=default_backend()
        )
        sec1_pem = priv.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.TraditionalOpenSSL,
            serialization.NoEncryption()
        ).decode()
        key_raw = ''.join(
            l for l in sec1_pem.strip().split('\n')
            if not l.startswith('-----')
        )

        cert_path = os.path.join(
            self.sdk_dir, "Data", "Certificates", "active_cert.cer"
        )
        key_path = os.path.join(
            self.sdk_dir, "Data", "Certificates", "active_key.pem"
        )
        open(cert_path, 'w').write(cert_raw)
        open(key_path,  'w').write(key_raw)

        log.debug("SDK cert written: %d chars", len(cert_raw))
        log.debug("SDK key written: %d chars", len(key_raw))
        return cert_path, key_path

    def _embed_qr(self, xml_str: str, qr_code: str) -> str:
        """Embed QR code into the XML before signing."""
        tree = etree.fromstring(xml_str.encode())
        for adr in tree.findall(f"{{{_CAC}}}AdditionalDocumentReference"):
            id_el = adr.find(f"{{{_CBC}}}ID")
            if id_el is not None and id_el.text == "QR":
                att = adr.find(f"{{{_CAC}}}Attachment")
                if att is None:
                    att = etree.SubElement(adr, f"{{{_CAC}}}Attachment")
                emb = att.find(f"{{{_CBC}}}EmbeddedDocumentBinaryObject")
                if emb is None:
                    emb = etree.SubElement(
                        att, f"{{{_CBC}}}EmbeddedDocumentBinaryObject"
                    )
                    emb.set("mimeCode", "text/plain")
                emb.text = qr_code
                break
        return etree.tostring(
            tree, pretty_print=True, xml_declaration=True, encoding="UTF-8"
        ).decode()

    def sign(self, xml_str: str, qr_code: str) -> tuple[str, str]:
        """
        Sign XML using ZATCA SDK.
        Returns (signed_xml, invoice_hash).
        """
        # Prepare paths
        input_path  = os.path.join(self.sdk_dir, "Data", "Input", "invoice_to_sign.xml")
        output_path = os.path.join(self.sdk_dir, "Data", "Input", "invoice_signed.xml")

        # QR already embedded upstream by _embed_qr_into_xml — write directly
        open(input_path, 'w', encoding='utf-8').write(xml_str)

        # Remove old output to detect failures
        if os.path.exists(output_path):
            os.remove(output_path)

        # Write cert + key for SDK
        cert_path, key_path = self._write_cert_and_key()

        # Update SDK config.json
        config_path = os.path.join(self.sdk_dir, "Configuration", "config.json")
        try:
            config = json.load(open(config_path))
        except Exception:
            config = {}
        config['certPath']       = cert_path
        config['privateKeyPath'] = key_path
        json.dump(config, open(config_path, 'w'), indent=2)

        # Run SDK signing
        cmd = [
            "java", "-jar", self.sdk_jar,
            "fatoora", "-sign",
            "-invoice", input_path,
            "-signedInvoice", output_path,
        ]
        log.info("Running ZATCA SDK sign...")
        r = subprocess.run(
            cmd, capture_output=True, text=True,
            cwd=self.sdk_dir, timeout=60
        )

        log.debug("SDK stdout: %s", r.stdout[-500:])
        if r.stderr:
            log.debug("SDK stderr: %s", r.stderr[-200:])

        if r.returncode != 0 or not os.path.exists(output_path):
            raise ZATCAError(
                f"SDK signing failed (rc={r.returncode}):\n"
                f"stdout: {r.stdout[-500:]}\n"
                f"stderr: {r.stderr[-300:]}"
            )

        # Extract invoice hash from SDK output
        match = _re.search(
            r'INVOICE HASH\s*=\s*([A-Za-z0-9+/=]+)', r.stdout
        )
        if not match:
            raise ZATCAError(
                f"Could not extract invoice hash from SDK output:\n{r.stdout}"
            )

        sdk_hash   = match.group(1)
        signed_xml = open(output_path, encoding='utf-8').read()

        log.info("SDK signed successfully. Hash: %s", sdk_hash)
        return signed_xml, sdk_hash

    def is_available(self) -> bool:
        """Check if SDK is available on this machine."""
        return os.path.exists(self.sdk_jar)


# ──────────────────────────────────────────────────────────
# Python XAdES-BES Signer (fallback — less reliable)
# ──────────────────────────────────────────────────────────

class ZATCASigner:
    """
    Python-based XAdES-BES signer.
    Used as fallback when SDK is not available.
    Note: ZATCA has rejected Python-signed invoices due to C14N differences.
    Prefer ZATCASDKSigner for production.
    """

    def __init__(self, private_key_pem: str, certificate_pem: str):
        self._priv = self._load_private_key(private_key_pem)
        cert_obj       = cx509.load_pem_x509_certificate(
            certificate_pem.encode(), default_backend()
        )
        self._cert_der = cert_obj.public_bytes(serialization.Encoding.DER)
        self._cert_b64 = base64.b64encode(self._cert_der).decode()
        self._cert_obj = cert_obj

    @staticmethod
    def _load_private_key(pem: str):
        import tempfile
        pem_bytes = pem.encode()
        try:
            return serialization.load_pem_private_key(
                pem_bytes, password=None, backend=default_backend()
            )
        except Exception:
            pass

        with tempfile.NamedTemporaryFile(suffix='.pem', mode='wb', delete=False) as f:
            f.write(pem_bytes)
            tmp_in = f.name
        tmp_out = tmp_in + '_p8.pem'
        try:
            r = subprocess.run(
                ['openssl', 'pkcs8', '-topk8', '-nocrypt',
                 '-in', tmp_in, '-out', tmp_out],
                capture_output=True, timeout=10,
            )
            if r.returncode == 0 and os.path.exists(tmp_out):
                pkcs8_bytes = open(tmp_out, 'rb').read()
                try:
                    return serialization.load_pem_private_key(
                        pkcs8_bytes, password=None, backend=default_backend()
                    )
                except Exception:
                    pass
        except Exception:
            pass
        finally:
            for p in (tmp_in, tmp_out):
                try: os.unlink(p)
                except Exception: pass

        raise ZATCAError(
            "Cannot load private key. Key must be secp256k1 in PKCS8 or SEC1 PEM format."
        )

    @staticmethod
    def _sha256_b64(data: bytes) -> str:
        return base64.b64encode(hashlib.sha256(data).digest()).decode()

    @staticmethod
    def _c14n(element_or_tree) -> bytes:
        buf = BytesIO()
        if isinstance(element_or_tree, etree._Element):
            etree.ElementTree(element_or_tree).write_c14n(
                buf, exclusive=False, with_comments=False
            )
        else:
            element_or_tree.write_c14n(buf, exclusive=False, with_comments=False)
        return buf.getvalue()

    def _ecdsa_raw_sign(self, data: bytes) -> bytes:
        return self._priv.sign(data, ec.ECDSA(hashes.SHA256()))

    def invoice_hash(self, xml_string: str) -> str:
        _EXT_NS = "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2"
        _CAC_NS = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
        _CBC_NS = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"

        tree_root = etree.fromstring(xml_string.encode("utf-8"))
        for ext in tree_root.findall(f"{{{_EXT_NS}}}UBLExtensions"):
            tree_root.remove(ext)
        for adr in tree_root.findall(f"{{{_CAC_NS}}}AdditionalDocumentReference"):
            id_el = adr.find(f"{{{_CBC_NS}}}ID")
            if id_el is not None and id_el.text == "QR":
                att = adr.find(f"{{{_CAC_NS}}}Attachment")
                if att is not None:
                    emb = att.find(f"{{{_CBC_NS}}}EmbeddedDocumentBinaryObject")
                    if emb is not None:
                        emb.text = ""
        canonical = self._c14n(tree_root)
        return base64.b64encode(hashlib.sha256(canonical).digest()).decode()

    def sign(self, xml_string: str, qr_code: str) -> tuple[str, str]:
        import copy as _copy
        tree = etree.fromstring(xml_string.encode())

        # Step 1: Embed QR
        for adr in tree.findall(f"{{{_CAC}}}AdditionalDocumentReference"):
            id_el = adr.find(f"{{{_CBC}}}ID")
            if id_el is not None and id_el.text == "QR":
                att = adr.find(f"{{{_CAC}}}Attachment")
                if att is None:
                    att = etree.SubElement(adr, f"{{{_CAC}}}Attachment")
                emb = att.find(f"{{{_CBC}}}EmbeddedDocumentBinaryObject")
                if emb is None:
                    emb = etree.SubElement(att, f"{{{_CBC}}}EmbeddedDocumentBinaryObject")
                    emb.set("mimeCode", "text/plain")
                emb.text = qr_code
                break

        signing_time = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        st_el = tree.find(f".//{{{_XADES}}}SigningTime")
        if st_el is not None:
            st_el.text = signing_time

        cert_digest = self._sha256_b64(self._cert_der)
        cdv = tree.find(f".//{{{_XADES}}}CertDigest/{{{_DS}}}DigestValue")
        if cdv is not None:
            cdv.text = cert_digest

        try:
            issuer_name   = self._cert_obj.issuer.rfc4514_string()
            serial_number = str(self._cert_obj.serial_number)
        except Exception:
            issuer_name   = ""
            serial_number = "0"

        iss_name_el = tree.find(f".//{{{_XADES}}}IssuerSerial/{{{_DS}}}X509IssuerName")
        if iss_name_el is not None:
            iss_name_el.text = issuer_name

        iss_sn_el = tree.find(f".//{{{_XADES}}}IssuerSerial/{{{_DS}}}X509SerialNumber")
        if iss_sn_el is not None:
            iss_sn_el.text = serial_number

        x5c = tree.find(f".//{{{_DS}}}X509Certificate")
        if x5c is not None:
            x5c.text = self._cert_b64

        _EXT_NS = "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2"
        hash_tree = _copy.deepcopy(tree)
        for ext in hash_tree.findall(f"{{{_EXT_NS}}}UBLExtensions"):
            hash_tree.remove(ext)
        for adr in hash_tree.findall(f"{{{_CAC}}}AdditionalDocumentReference"):
            id_el = adr.find(f"{{{_CBC}}}ID")
            if id_el is not None and id_el.text == "QR":
                att = adr.find(f"{{{_CAC}}}Attachment")
                if att is not None:
                    emb = att.find(f"{{{_CBC}}}EmbeddedDocumentBinaryObject")
                    if emb is not None:
                        emb.text = ""
                break

        inv_digest = self._sha256_b64(self._c14n(hash_tree))
        refs = tree.findall(f".//{{{_DS}}}Reference")
        for ref_el in refs:
            if ref_el.get("Id") == "id-doc":
                dv = ref_el.find(f"{{{_DS}}}DigestValue")
                if dv is not None:
                    dv.text = inv_digest
                break

        sp_el = tree.find(f".//{{{_XADES}}}SignedProperties")
        if sp_el is not None:
            sp_digest = self._sha256_b64(self._c14n(sp_el))
            for ref_el in refs:
                if ref_el.get("Id") == "id-sp":
                    dv = ref_el.find(f"{{{_DS}}}DigestValue")
                    if dv is not None:
                        dv.text = sp_digest
                    break

        si_el = tree.find(f".//{{{_DS}}}SignedInfo")
        if si_el is None:
            raise ZATCAError("ds:SignedInfo not found")

        si_bytes = self._c14n(si_el)
        raw_sig  = self._ecdsa_raw_sign(si_bytes)
        sig_b64  = base64.b64encode(raw_sig).decode()

        sv_el = tree.find(f".//{{{_DS}}}SignatureValue")
        if sv_el is not None:
            sv_el.text = sig_b64

        signed_xml = etree.tostring(
            tree, pretty_print=True, xml_declaration=True, encoding="UTF-8"
        ).decode()
        inv_hash = self.invoice_hash(signed_xml)
        return signed_xml, inv_hash


# ──────────────────────────────────────────────────────────
# UBL 2.1 XML Generator
# ──────────────────────────────────────────────────────────

class ZATCAXMLGenerator:
    """
    Builds ZATCA Phase-2 UBL 2.1 XML.

    UBL 2.1 element order (enforced by XSD):
      UBLExtensions → ProfileID → ID → UUID → IssueDate → IssueTime →
      InvoiceTypeCode → DocumentCurrencyCode → TaxCurrencyCode →
      InvoicePeriod → BillingReference → AdditionalDocumentReference →
      AccountingSupplierParty → AccountingCustomerParty →
      Delivery → PaymentMeans → TaxTotal → LegalMonetaryTotal → InvoiceLine

    ZATCA BT mapping:
      StreetName           = BT-35 (seller) / BT-50 (buyer)
      AdditionalStreetName = KSA-3 (seller district) / KSA-4 (buyer district)
      BuildingNumber       = KSA-17 (seller) / KSA-18 (buyer) [4 digits]
      CityName             = BT-37 (seller) / BT-52 (buyer)
      PostalZone           = BT-38 (seller) / BT-53 (buyer) [5 digits]
    """

    INVOICE_TYPES = {
        "invoice":    ("388", "0100000"),
        "simplified": ("388", "0200000"),
        "credit":     ("381", "0100000"),
        "debit":      ("383", "0100000"),
    }

    @staticmethod
    def _sub(parent, ns, local, text=None, **attribs):
        el = etree.SubElement(parent, f"{{{ns}}}{local}")
        for k, v in attribs.items():
            el.set(k, str(v))
        if text is not None:
            el.text = str(text)
        return el

    def _cbc(self, p, l, t=None, **a): return self._sub(p, _CBC, l, t, **a)
    def _cac(self, p, l):             return self._sub(p, _CAC, l)
    def _ds(self,  p, l, t=None, **a): return self._sub(p, _DS,  l, t, **a)
    def _xd(self,  p, l, t=None, **a): return self._sub(p, _XADES, l, t, **a)

    def generate(self, invoice) -> str:
        root = etree.Element(f"{{{NS[None]}}}Invoice", nsmap=NS)
        self._ubl_extensions(root)
        self._header(root, invoice)
        self._additional_docs(root, invoice)
        self._party_section(root, "AccountingSupplierParty", invoice.company, True)
        self._party_section(root, "AccountingCustomerParty", invoice.client,  False)
        self._delivery(root, invoice)       # KSA-5: after party sections
        self._payment_means(root, invoice)
        self._tax_total(root, invoice)
        self._monetary_total(root, invoice)
        self._lines(root, invoice)
        return etree.tostring(
            root, pretty_print=True, xml_declaration=True, encoding="UTF-8"
        ).decode()

    _SP_ID      = "urn:oasis:names:specification:ubl:signature:1"
    _SIG_ID     = "urn:oasis:names:specification:ubl:signature:Invoice"
    _SP_ELEM_ID = "xades-signed-props"

    def _ubl_extensions(self, root):
        ue  = self._sub(root, _EXT, "UBLExtensions")
        e   = self._sub(ue,   _EXT, "UBLExtension")
        uri = self._sub(e,    _EXT, "ExtensionURI")
        uri.text = "urn:oasis:names:specification:ubl:dsig:enveloped:xades"
        ec_ = self._sub(e, _EXT, "ExtensionContent")

        sig = self._sub(ec_, _DS, "Signature")
        si  = self._ds(sig, "SignedInfo")
        self._ds(si, "CanonicalizationMethod",
                 Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315")
        self._ds(si, "SignatureMethod",
                 Algorithm="http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256")

        ref_doc = self._ds(si, "Reference", Id="id-doc", URI="")
        xf = self._ds(ref_doc, "Transforms")
        self._ds(xf, "Transform",
                 Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315")
        self._ds(ref_doc, "DigestMethod",
                 Algorithm="http://www.w3.org/2001/04/xmlenc#sha256")
        self._ds(ref_doc, "DigestValue", "")

        ref_sp = self._ds(si, "Reference",
                          Id="id-sp",
                          Type="http://uri.etsi.org/01903#SignedProperties",
                          URI=f"#{self._SP_ID}")
        self._ds(ref_sp, "DigestMethod",
                 Algorithm="http://www.w3.org/2001/04/xmlenc#sha256")
        self._ds(ref_sp, "DigestValue", "")

        self._ds(sig, "SignatureValue", "")
        ki  = self._ds(sig, "KeyInfo")
        x5d = self._ds(ki,  "X509Data")
        self._ds(x5d, "X509Certificate", "")

        obj = self._ds(sig, "Object")
        qp  = self._xd(obj, "QualifyingProperties", Target=self._SIG_ID)
        sp  = self._xd(qp, "SignedProperties", Id=self._SP_ELEM_ID)
        ssp = self._xd(sp, "SignedSignatureProperties")
        st  = self._xd(ssp, "SigningTime")
        st.text = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        sc      = self._xd(ssp, "SigningCertificate")
        cert_el = self._xd(sc,  "Cert")
        cd      = self._xd(cert_el, "CertDigest")
        self._ds(cd, "DigestMethod",
                 Algorithm="http://www.w3.org/2001/04/xmlenc#sha256")
        self._ds(cd, "DigestValue", "")

        iss = self._xd(cert_el, "IssuerSerial")
        self._ds(iss, "X509IssuerName",   "")
        self._ds(iss, "X509SerialNumber", "0")

    def _header(self, root, invoice):
        type_code, subtype = self.INVOICE_TYPES.get(
            invoice.transaction_type, ("388", "0100000"))

        self._cbc(root, "ProfileID",    "reporting:1.0")
        self._cbc(root, "ID",           str(invoice.invoice_number))
        self._cbc(root, "UUID",         str(invoice.invoice_uuid))
        self._cbc(root, "IssueDate",    invoice.date.strftime("%Y-%m-%d"))
        self._cbc(root, "IssueTime",    "00:00:00")

        itc = self._cbc(root, "InvoiceTypeCode", type_code)
        itc.set("name", subtype)

        self._cbc(root, "DocumentCurrencyCode", "SAR")
        self._cbc(root, "TaxCurrencyCode",      "SAR")

        is_standard = invoice.transaction_type in ("invoice", "credit", "debit")
        if is_standard:
            sd = getattr(invoice, "supply_date", None) or invoice.date
            period = self._cac(root, "InvoicePeriod")
            self._cbc(period, "StartDate", sd.strftime("%Y-%m-%d"))
            self._cbc(period, "EndDate",   sd.strftime("%Y-%m-%d"))

        if invoice.transaction_type in ("credit", "debit"):
            ref_inv = getattr(invoice, "reference_invoice", None)
            if ref_inv:
                br  = self._cac(root, "BillingReference")
                idr = self._cac(br,  "InvoiceDocumentReference")
                self._cbc(idr, "ID", str(ref_inv.invoice_number))

    def _additional_docs(self, root, invoice):
        # ICV — KSA-16
        icv = self._cac(root, "AdditionalDocumentReference")
        self._cbc(icv, "ID", "ICV")
        self._cbc(icv, "UUID", str(getattr(invoice, "invoice_counter_value", 1)))

        # PIH — KSA-13
        pih_hash = (
            getattr(invoice, "previous_invoice_hash", None) or ZATCA_GENESIS_HASH
        )
        pih = self._cac(root, "AdditionalDocumentReference")
        self._cbc(pih, "ID", "PIH")
        att = self._cac(pih, "Attachment")
        emb = self._cbc(att, "EmbeddedDocumentBinaryObject", pih_hash)
        emb.set("mimeCode", "text/plain")

        # QR — placeholder, SDK embeds the real value during signing
        qr = self._cac(root, "AdditionalDocumentReference")
        self._cbc(qr, "ID", "QR")
        qr_att = self._cac(qr, "Attachment")
        qr_emb = self._cbc(qr_att, "EmbeddedDocumentBinaryObject", "")
        qr_emb.set("mimeCode", "text/plain")

    def _party_section(self, root, tag, entity, is_supplier):
        container = self._cac(root, tag)
        party = self._cac(container, "Party")

        cr  = (getattr(entity, "cr_number",  "") or "").strip()
        vat = (getattr(entity, "vat_number", "") or "").strip()

        if cr:
            pid = self._cac(party, "PartyIdentification")
            el  = self._cbc(pid, "ID", cr)
            el.set("schemeID", "CRN")

        addr = self._cac(party, "PostalAddress")

        # StreetName — required, no fallback
        street = (getattr(entity, "street_name", "") or "").strip()
        if not street:
            raise ZATCAValidationError(
                f"اسم الشارع مطلوب لـ {getattr(entity, 'name', tag)}"
            )
        self._cbc(addr, "StreetName", street[:200])

        # AdditionalStreetName — optional
        additional_street = (getattr(entity, "additional_street", "") or "").strip()
        if additional_street:
            self._cbc(addr, "AdditionalStreetName", additional_street[:127])

        # BuildingNumber — must be 4 digits, no "0000" fallback
        bld_raw = str(getattr(entity, "building_number", "") or "").strip()
        digits  = ''.join(c for c in bld_raw if c.isdigit())
        if not digits:
            raise ZATCAValidationError(
                f"رقم المبنى مطلوب (4 أرقام) لـ {getattr(entity, 'name', tag)}"
            )
        bld = digits.zfill(4)[:4]
        self._cbc(addr, "BuildingNumber", bld)

        # CitySubdivisionName = district — required, no "غير محدد" fallback
        district = (getattr(entity, "district", "") or "").strip()
        if not district:
            raise ZATCAValidationError(
                f"الحي مطلوب لـ {getattr(entity, 'name', tag)}"
            )
        self._cbc(addr, "CitySubdivisionName", district[:127])

        # CityName — required
        city = (getattr(entity, "city", "") or "الرياض").strip()
        self._cbc(addr, "CityName", city[:127])

        # PostalZone — 5 digits, no "00000" fallback
        pz_raw = str(getattr(entity, "postal_code", "") or "").strip()
        pz_dig = ''.join(c for c in pz_raw if c.isdigit())
        if not pz_dig:
            raise ZATCAValidationError(
                f"الرمز البريدي مطلوب (5 أرقام) لـ {getattr(entity, 'name', tag)}"
            )
        pz = pz_dig.zfill(5)[:5]
        self._cbc(addr, "PostalZone", pz)

        # Country
        cntry = self._cac(addr, "Country")
        self._cbc(cntry, "IdentificationCode", "SA")

        # Legal Entity
        if vat:
            pts = self._cac(party, "PartyTaxScheme")
            self._cbc(pts, "CompanyID", vat)
            ts  = self._cac(pts, "TaxScheme")
            self._cbc(ts,  "ID", "VAT")

        ple = self._cac(party, "PartyLegalEntity")
        self._cbc(ple, "RegistrationName", entity.name[:200])

    def _delivery(self, root, invoice):
        """KSA-5: ActualDeliveryDate — required for standard tax invoices."""
        is_standard = invoice.transaction_type in ("invoice", "credit", "debit")
        if is_standard:
            sd = getattr(invoice, "supply_date", None) or invoice.date
            delivery = self._cac(root, "Delivery")
            self._cbc(delivery, "ActualDeliveryDate", sd.strftime("%Y-%m-%d"))

    def _payment_means(self, root, invoice):
        MAP = {"cash": "10", "bank": "42", "cheque": "20",
               "card": "48", "credit": "30"}
        pm = self._cac(root, "PaymentMeans")
        self._cbc(pm, "PaymentMeansCode",
                  MAP.get(getattr(invoice, "payment_method", "credit"), "30"))

    def _tax_total(self, root, invoice):
        tax_str = f"{float(invoice.total_tax):.2f}"
        txb_str = f"{float(invoice.taxable_amount):.2f}"
        pct_str = f"{float(invoice.tax_rate):.2f}"

        tt1 = self._cac(root, "TaxTotal")
        self._cbc(tt1, "TaxAmount", tax_str, currencyID="SAR")

        tt2 = self._cac(root, "TaxTotal")
        self._cbc(tt2, "TaxAmount", tax_str, currencyID="SAR")
        sub = self._cac(tt2, "TaxSubtotal")
        self._cbc(sub, "TaxableAmount", txb_str, currencyID="SAR")
        self._cbc(sub, "TaxAmount",     tax_str, currencyID="SAR")
        tc  = self._cac(sub, "TaxCategory")
        self._cbc(tc, "ID",      "S")
        self._cbc(tc, "Percent", pct_str)
        ts  = self._cac(tc, "TaxScheme")
        self._cbc(ts, "ID", "VAT")

    def _monetary_total(self, root, invoice):
        lmt = self._cac(root, "LegalMonetaryTotal")
        self._cbc(lmt, "LineExtensionAmount",
                  f"{float(invoice.total_sales_excl_tax):.2f}", currencyID="SAR")
        self._cbc(lmt, "TaxExclusiveAmount",
                  f"{float(invoice.taxable_amount):.2f}", currencyID="SAR")
        self._cbc(lmt, "TaxInclusiveAmount",
                  f"{float(invoice.total):.2f}", currencyID="SAR")
        disc = float(getattr(invoice, "discount_amount", 0) or 0)
        if disc > 0:
            self._cbc(lmt, "AllowanceTotalAmount",
                      f"{disc:.2f}", currencyID="SAR")
        self._cbc(lmt, "PayableAmount",
                  f"{float(invoice.total):.2f}", currencyID="SAR")

    def _lines(self, root, invoice):
        for idx, item in enumerate(invoice.items.all(), start=1):
            line = self._cac(root, "InvoiceLine")
            self._cbc(line, "ID", str(idx))
            self._cbc(line, "InvoicedQuantity",
                      f"{float(item.quantity):.3f}",
                      unitCode=getattr(item, "unit", None) or "PCE")

            net     = float(item.total)
            line_vat = float(item.tax_amount)

            self._cbc(line, "LineExtensionAmount",
                      f"{net:.2f}", currencyID="SAR")

            tt = self._cac(line, "TaxTotal")
            self._cbc(tt, "TaxAmount",      f"{line_vat:.2f}", currencyID="SAR")
            self._cbc(tt, "RoundingAmount", f"{net + line_vat:.2f}", currencyID="SAR")

            it = self._cac(line, "Item")
            self._cbc(it, "Name", str(item.description)[:100])
            tc  = self._cac(it, "ClassifiedTaxCategory")
            self._cbc(tc, "ID",      "S")
            self._cbc(tc, "Percent", f"{float(item.tax_rate):.2f}")
            ts  = self._cac(tc, "TaxScheme")
            self._cbc(ts, "ID", "VAT")

            price = self._cac(line, "Price")
            self._cbc(price, "PriceAmount",
                      f"{float(item.unit_price):.2f}", currencyID="SAR")
            self._cbc(price, "BaseQuantity",
                      "1", unitCode=getattr(item, "unit", None) or "PCE")


# ──────────────────────────────────────────────────────────
# QR Code Generator
# ──────────────────────────────────────────────────────────

class ZATCAQRGenerator:
    """
    5-field TLV QR code (ZATCA BR-KSA-27).
    Output MUST be <= 1000 chars (BR-CL-KSA-14).
    """

    _NAME_MAX_BYTES = 50

    def _tlv(self, tag: int, value: str) -> bytes:
        data = value.encode("utf-8")
        if len(data) > 127:
            data = data[:127]
        return bytes([tag, len(data)]) + data

    def generate(self, invoice, signing_time: str = None) -> str:
        company = invoice.company
        vat = (getattr(company, "vat_number", "") or "").strip()

        if not vat or len(vat) != 15 or not vat.isdigit():
            raise ZATCAValidationError(
                "الرقم الضريبي للشركة يجب أن يكون 15 رقماً صحيحاً"
            )

        raw_name = (getattr(company, "name", "") or "").strip()
        if not raw_name:
            raise ZATCAValidationError("اسم البائع مطلوب للـ QR Code")

        # Tighter name truncation — stay well under 1000 char limit
        # Total TLV overhead: tag+len bytes + vat(15) + timestamp(20) + total(~10) + vat_amount(~10)
        # Reserve 700 chars for base64 output → name can be at most 30 bytes safely
        name_bytes = raw_name.encode("utf-8")
        if len(name_bytes) > 30:
            name_bytes = name_bytes[:30]
            raw_name   = name_bytes.decode("utf-8", errors="ignore")

        ts     = signing_time or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        total  = f"{float(invoice.total):.2f}"
        vat_am = f"{float(invoice.total_tax):.2f}"

        tlv = (
            self._tlv(1, raw_name) +
            self._tlv(2, vat)      +
            self._tlv(3, ts)       +
            self._tlv(4, total)    +
            self._tlv(5, vat_am)
        )
        result = base64.b64encode(tlv).decode()

        if len(result) > 1000:
            raise ZATCAValidationError(
                f"QR base64 length {len(result)} exceeds 1000 chars — shorten company name"
            )
        return result

    def generate_png(self, invoice) -> bytes:
        qr_b64 = self.generate(invoice)
        img    = qrcode.make(qr_b64)
        buf    = BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()


# ──────────────────────────────────────────────────────────
# ZATCA API Client
# ──────────────────────────────────────────────────────────

class ZATCAAPIClient:

    def __init__(self, csid: str, secret: str, environment: str = "sandbox"):
        self.csid    = csid
        self.secret  = secret
        if environment == "production":
            self.base_url = ZATCA_PROD_BASE
        elif environment == "simulation":
            self.base_url = ZATCA_SIMULATION_BASE
        else:
            self.base_url = ZATCA_SANDBOX_BASE
        self.environment = environment

    def _auth(self) -> dict:
        token = base64.b64encode(f"{self.csid}:{self.secret}".encode()).decode()
        return {
            "Authorization":   f"Basic {token}",
            "Content-Type":    "application/json",
            "accept":          "application/json",
            "Accept-Version":  "V2",
            "Accept-Language": "en",
        }

    def _payload(self, signed_xml, invoice_hash, uuid_val) -> dict:
        return {
            "invoiceHash": invoice_hash,
            "uuid":        str(uuid_val),
            "invoice":     base64.b64encode(signed_xml.encode()).decode(),
        }

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=4, max=10),
        retry=retry_if_exception_type((requests.RequestException, ConnectionError)),
        reraise=True,
    )
    def _post(self, url, payload, extra_headers=None) -> requests.Response:
        headers = self._auth()
        if extra_headers:
            headers.update(extra_headers)
        return requests.post(url, json=payload, headers=headers, timeout=30)

    def submit_clearance(self, invoice, signed_xml, invoice_hash) -> dict:
        payload = self._payload(signed_xml, invoice_hash, invoice.invoice_uuid)
        resp    = self._post(
            f"{self.base_url}{ENDPOINTS['clearance']}", payload,
            extra_headers={"Clearance-Status": "CLEARED"},
        )
        data = resp.json() if resp.content else {}
        return {
            "status_code":     resp.status_code,
            "success":         resp.status_code in (200, 202),
            "data":            data,
            "cleared_invoice": data.get("clearedInvoice"),
        }

    def submit_reporting(self, invoice, signed_xml, invoice_hash) -> dict:
        payload = self._payload(signed_xml, invoice_hash, invoice.invoice_uuid)
        resp    = self._post(
            f"{self.base_url}{ENDPOINTS['reporting']}", payload,
            extra_headers={"Clearance-Status": "REPORTED"},
        )
        data = resp.json() if resp.content else {}
        return {
            "status_code": resp.status_code,
            "success":     resp.status_code in (200, 202),
            "data":        data,
        }

    def compliance_check(self, invoice, signed_xml, invoice_hash) -> dict:
        payload = self._payload(signed_xml, invoice_hash, invoice.invoice_uuid)
        resp    = self._post(
            f"{self.base_url}{ENDPOINTS['compliance_invoices']}", payload
        )
        data = resp.json() if resp.content else {}
        return {
            "status_code": resp.status_code,
            "success":     resp.status_code in (200, 202),
            "data":        data,
        }


# ──────────────────────────────────────────────────────────
# Pre-submission Validation
# ──────────────────────────────────────────────────────────

def validate_invoice(invoice) -> None:
    """
    Validate invoice data before ZATCA submission.
    Raises ZATCAValidationError with all errors if any are found.
    """
    errors = []

    # ── Seller (Company) ────────────────────────────────────
    co = getattr(invoice, "company", None)
    if not co:
        errors.append("بيانات الشركة مفقودة")
    else:
        vat = (getattr(co, "vat_number", "") or "").strip()
        if not vat or len(vat) != 15 or not vat.isdigit():
            errors.append("الرقم الضريبي للشركة يجب أن يكون 15 رقماً صحيحاً")

        # All seller address fields required (BR-KSA-09)
        for field, label in [
            ("street_name",     "اسم الشارع"),
            ("building_number", "رقم المبنى"),
            ("postal_code",     "الرمز البريدي"),
            ("city",            "المدينة"),
            ("district",        "الحي"),
        ]:
            val = (getattr(co, field, "") or "").strip()
            if not val:
                errors.append(f"{label} للشركة مطلوب (BR-KSA-09)")

    # ── Buyer (Client) ──────────────────────────────────────
    cl = getattr(invoice, "client", None)
    if cl:
        # If buyer is in SA, address fields are mandatory (BR-KSA-63)
        country = (getattr(cl, "country", "") or "Saudi Arabia").strip().lower()
        is_saudi = "saudi" in country or country == "sa"

        if is_saudi:
            for field, label in [
                ("street_name",     "اسم شارع العميل"),
                ("building_number", "رقم مبنى العميل"),
                ("postal_code",     "الرمز البريدي للعميل"),
                ("city",            "مدينة العميل"),
                ("district",        "حي العميل"),
            ]:
                val = (getattr(cl, field, "") or "").strip()
                if not val:
                    errors.append(f"{label} مطلوب (BR-KSA-63)")

    # ── Invoice ─────────────────────────────────────────────
    if not getattr(invoice, "invoice_uuid", None):
        errors.append("UUID الفاتورة مطلوب")

    if float(getattr(invoice, "total", 0)) <= 0:
        errors.append("إجمالي الفاتورة يجب أن يكون أكبر من صفر")

    if invoice.items.count() == 0:
        errors.append("الفاتورة يجب أن تحتوي على بند واحد على الأقل")

    # Supply date required for standard invoices (BR-KSA-15)
    if invoice.transaction_type in ("invoice", "credit", "debit"):
        if not getattr(invoice, "supply_date", None):
            errors.append("تاريخ التوريد مطلوب للفواتير الضريبية (BR-KSA-15)")

    if errors:
        raise ZATCAValidationError(
            "فشل التحقق قبل الإرسال:\n" + "\n".join(f"  • {e}" for e in errors)
        )


# ──────────────────────────────────────────────────────────
# High-level Orchestration
# ──────────────────────────────────────────────────────────

class ZATCAInvoiceService:
    """
    Full pipeline:
      validate → assign counter → chain hash → QR → XML →
      sign (SDK first, Python fallback) → submit → log.

    Usage:
        device = ZATCADevice.objects.get(company=inv.company, is_active=True)
        result = ZATCAInvoiceService(device).process(inv)
    """

    def __init__(self, device):
        self.device    = device
        self.generator = ZATCAXMLGenerator()
        self.qr_gen    = ZATCAQRGenerator()
        self.signer    = ZATCASigner(device.private_key, device.certificate)
        self.sdk_signer = ZATCASDKSigner(device)
        self.client    = ZATCAAPIClient(device.csid, device.secret, device.environment)

    def _assign_counter_and_hash(self, invoice) -> None:
        from django.db import transaction as db_transaction

        with db_transaction.atomic():
            locked = type(invoice).objects.select_for_update().get(pk=invoice.pk)

            if not locked.invoice_counter_value or locked.invoice_counter_value < 1:
                last = (
                    type(invoice).objects
                    .filter(transaction_type=invoice.transaction_type)
                    .exclude(pk=invoice.pk)
                    .filter(invoice_counter_value__gte=1)
                    .select_for_update()
                    .order_by('-invoice_counter_value')
                    .first()
                )
                locked.invoice_counter_value = (
                    last.invoice_counter_value + 1 if last else 1
                )

            if not locked.previous_invoice_hash:
                prev = (
                    type(invoice).objects
                    .filter(
                        transaction_type=invoice.transaction_type,
                        zatca_invoice_hash__isnull=False,
                    )
                    .exclude(zatca_invoice_hash='')
                    .exclude(pk=invoice.pk)
                    .order_by('-invoice_counter_value')
                    .first()
                )
                locked.previous_invoice_hash = (
                    prev.zatca_invoice_hash if prev else ZATCA_GENESIS_HASH
                )

            locked.save(update_fields=['invoice_counter_value', 'previous_invoice_hash'])
            invoice.invoice_counter_value = locked.invoice_counter_value
            invoice.previous_invoice_hash = locked.previous_invoice_hash
            
    def _embed_qr_into_xml(self, xml_str: str, qr_code: str) -> str:
        """Embed QR code into XML before passing to SDK signer."""
        from lxml import etree
        _CAC_NS = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
        _CBC_NS = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"

        tree = etree.fromstring(xml_str.encode())
        for adr in tree.findall(f"{{{_CAC_NS}}}AdditionalDocumentReference"):
            id_el = adr.find(f"{{{_CBC_NS}}}ID")
            if id_el is not None and id_el.text == "QR":
                att = adr.find(f"{{{_CAC_NS}}}Attachment")
                if att is None:
                    att = etree.SubElement(adr, f"{{{_CAC_NS}}}Attachment")
                emb = att.find(f"{{{_CBC_NS}}}EmbeddedDocumentBinaryObject")
                if emb is None:
                    emb = etree.SubElement(att, f"{{{_CBC_NS}}}EmbeddedDocumentBinaryObject")
                    emb.set("mimeCode", "text/plain")
                emb.text = qr_code
                break
        return etree.tostring(
            tree, pretty_print=True, xml_declaration=True, encoding="UTF-8"
        ).decode()
        
    def _sign_invoice(self, xml_str: str, qr_code: str) -> tuple[str, str]:
        """
        Sign using SDK (preferred) with Python fallback.
        Returns (signed_xml, invoice_hash).
        """
        if self.sdk_signer.is_available():
            try:
                log.info("Signing with ZATCA SDK...")
                return self.sdk_signer.sign(xml_str, qr_code)
            except Exception as e:
                log.warning("SDK signing failed, falling back to Python: %s", e)

        log.warning("Using Python signer (SDK not available or failed)")
        return self.signer.sign(xml_str, qr_code)

    def process(self, invoice) -> dict:
        from .models import ZATCALog
        from django.utils import timezone as tz

        try:
            # Pre-flight validation
            validate_invoice(invoice)

            # Atomically assign counter + chain hash
            self._assign_counter_and_hash(invoice)

            # Generate QR code
            qr_code = self.qr_gen.generate(invoice)
            invoice.zatca_qr_code = qr_code
            invoice.save(update_fields=["zatca_qr_code"])

            # Generate XML
            xml_str = self.generator.generate(invoice)

            # Embed QR into XML before signing — SDK requires it present
            xml_str = self._embed_qr_into_xml(xml_str, qr_code)

            # Sign (SDK preferred, Python fallback)
            signed_xml, inv_hash = self._sign_invoice(xml_str, qr_code)

            # Determine submission endpoint
            is_b2b      = bool(getattr(invoice.client, "vat_number", None))
            is_standard = invoice.transaction_type in ("invoice", "credit", "debit")

            if is_b2b or is_standard:
                result = self.client.submit_clearance(invoice, signed_xml, inv_hash)
                action = "clearance"
            else:
                result = self.client.submit_reporting(invoice, signed_xml, inv_hash)
                action = "reporting"

            success = result.get("success", False)

            # Use ZATCA-cleared XML if returned
            if success and action == "clearance" and result.get("cleared_invoice"):
                try:
                    signed_xml = base64.b64decode(
                        result["cleared_invoice"]
                    ).decode("utf-8")
                except Exception:
                    pass

            # Persist results
            invoice.zatca_xml           = xml_str
            invoice.zatca_signed_xml    = signed_xml
            invoice.zatca_invoice_hash  = inv_hash
            invoice.zatca_qr_code       = qr_code
            invoice.zatca_response      = result.get("data")
            invoice.zatca_submitted_at  = tz.now()
            invoice.zatca_submission_id = result.get("data", {}).get("uuid", "")
            invoice.zatca_status = (
                ("cleared" if action == "clearance" else "reported")
                if success else "rejected"
            )
            invoice.save(update_fields=[
                "zatca_xml", "zatca_signed_xml", "zatca_invoice_hash",
                "zatca_qr_code", "zatca_response", "zatca_submitted_at",
                "zatca_submission_id", "zatca_status",
            ])

            # Audit log
            ZATCALog.objects.create(
                device=self.device,
                sales_invoice=invoice,
                action=action,
                request_payload=base64.b64encode(
                    signed_xml.encode()
                ).decode()[:51200],
                response_code=result.get("status_code"),
                response_body=json.dumps(result.get("data", {})),
                success=success,
                error_message="" if success else json.dumps(result.get("data", {})),
            )

            return {
                "success":      success,
                "result":       result,
                "qr_code":      qr_code,
                "invoice_hash": inv_hash,
            }

        except ZATCAValidationError:
            raise
        except Exception as exc:
            ZATCALog.objects.create(
                device=self.device,
                sales_invoice=invoice,
                action="clearance",
                success=False,
                error_message=str(exc),
            )
            raise


# ──────────────────────────────────────────────────────────
# Certificate utilities
# ──────────────────────────────────────────────────────────

def _der_b64_to_pem(der_b64: str) -> str:
    """Convert ZATCA binarySecurityToken to PEM certificate."""
    token = der_b64.strip()

    if token.startswith("-----BEGIN CERTIFICATE"):
        return token

    try:
        raw = base64.b64decode(token)
    except Exception:
        return token

    if len(raw) > 2 and raw[0] not in (0x30,):
        try:
            raw = base64.b64decode(raw)
        except Exception:
            pass

    try:
        cert = cx509.load_der_x509_certificate(raw, default_backend())
        return cert.public_bytes(serialization.Encoding.PEM).decode()
    except Exception:
        return (
            "-----BEGIN CERTIFICATE-----\n"
            + base64.b64encode(raw).decode()
            + "\n-----END CERTIFICATE-----\n"
        )
    
# ──────────────────────────────────────────────────────────
# Onboarding
# ──────────────────────────────────────────────────────────

def onboard_device(
    company,
    device_name: str,
    otp: str,
    environment: str = "sandbox",
    csr_pem: str = None,
    private_key_pem: str = None,
):
    """
    Full ZATCA onboarding. Call once with OTP from the Fatoora portal.
    Creates and returns a ZATCADevice instance.

    For simulation: uses compliance CSID directly (production CSID not supported).
    For production: gets compliance CSID then production CSID.

    Args:
        company:         Company model instance
        device_name:     Name for the EGS device (e.g. "EGS-001")
        otp:             6-digit OTP from the Fatoora portal
        environment:     "sandbox" | "simulation" | "production"
        csr_pem:         Optional pre-generated CSR PEM.
        private_key_pem: PKCS8 PEM of the private key matching the CSR.
    """
    from .models import ZATCADevice

    log.info("Starting ZATCA onboarding: company=%s env=%s", company.name, environment)

    # Step 1: Key pair + CSR
    if csr_pem and private_key_pem:
        priv_pem = private_key_pem
        serial   = f"{company.cr_number or 'EGS'}-{device_name}"
        log.info("Using pre-generated CSR (%d chars)", len(csr_pem))
    else:
        priv_pem, _ = generate_ec_keypair()
        serial       = f"{company.cr_number or 'EGS'}-{device_name}-{uuid.uuid4().hex[:8]}"
        csr_pem, _   = generate_csr(
            company_name=company.name_en or company.name,
            vat_number=company.vat_number,
            serial_number=serial,
            private_key_pem=priv_pem,
        )
        log.info("Auto-generated CSR for serial=%s", serial)

    # Step 2: Compliance CSID
    ob         = ZATCAOnboarding(environment)
    compliance = ob.get_compliance_csid(csr_pem, otp)
    if not compliance.get("csid"):
        raise ZATCAError(
            f"فشل الحصول على Compliance CSID: {compliance.get('dispositionMessage')}"
        )
    log.info("Compliance CSID obtained. requestID=%s", compliance.get("requestID"))

    # Step 3: Production CSID (production only)
    # Simulation environment does not support production CSID endpoint
    final_csid   = compliance["csid"]
    final_secret = compliance["secret"]
    cert_token   = compliance["csid"]  # compliance cert used for signing

    if environment == "production":
        try:
            production = ob.get_production_csid(
                compliance_csid=compliance["csid"],
                compliance_secret=compliance["secret"],
                compliance_request_id=compliance["requestID"],
            )
            if production.get("csid"):
                final_csid   = production["csid"]
                final_secret = production["secret"]
                log.info("Production CSID obtained.")
            else:
                log.warning("Production CSID response empty, using compliance CSID")
        except Exception as e:
            log.warning("Production CSID failed: %s — using compliance CSID", e)
    else:
        log.info("Environment=%s — using compliance CSID for submissions", environment)

    # Step 4: Deactivate old devices
    ZATCADevice.objects.filter(
        company=company, environment=environment
    ).update(is_active=False)

    # Step 5: Persist device
    device = ZATCADevice.objects.create(
        company=company,
        device_name=device_name,
        serial_number=serial,
        private_key=priv_pem,
        csr=csr_pem,
        certificate=_der_b64_to_pem(cert_token),
        csid=final_csid,
        secret=final_secret,
        environment=environment,
        is_active=True,
        registered_at=datetime.now(timezone.utc),
    )

    log.info("ZATCA onboarding complete. Device id=%s", device.id)
    return device