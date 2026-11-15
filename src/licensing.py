"""One-time online activation with offline certificate verification."""

import base64
import ctypes
from dataclasses import dataclass
from datetime import datetime
import json
import os
from pathlib import Path
from typing import Any, Callable, Dict, Optional, Tuple
import urllib.error
import urllib.request

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import (
    Ed25519PrivateKey,
    Ed25519PublicKey,
)


PRODUCT_NAME = "Inv_Reader"
LICENSE_FILENAME = "license.json"
DEVICE_KEY_FILENAME = "device.key"
ACTIVATION_PATH = "/v1/activate"


class ActivationError(RuntimeError):
    """The application cannot prove that this machine is licensed."""


def _b64encode(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


def _b64decode(value: str) -> bytes:
    if not isinstance(value, str):
        raise ActivationError("Activation data is malformed")
    try:
        return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))
    except (ValueError, TypeError) as error:
        raise ActivationError("Activation data is malformed") from error


def _canonical_json(value: Dict[str, Any]) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")


def _public_key_bytes(public_key: Ed25519PublicKey) -> bytes:
    return public_key.public_bytes(
        encoding=serialization.Encoding.Raw,
        format=serialization.PublicFormat.Raw,
    )


def create_device_identity() -> Tuple[Ed25519PrivateKey, str]:
    """Create a machine identity and return its private key and public key text."""
    private_key = Ed25519PrivateKey.generate()
    return private_key, _b64encode(_public_key_bytes(private_key.public_key()))


def create_certificate(server_private_key: Ed25519PrivateKey,
                       payload: Dict[str, Any]) -> Dict[str, str]:
    """Create a signed certificate that can be verified without the server."""
    certificate_payload = dict(payload)
    certificate_payload.setdefault("product", PRODUCT_NAME)
    certificate_payload.setdefault("schema", 1)
    certificate_payload.setdefault("issued_at", datetime.utcnow().isoformat() + "Z")
    encoded_payload = _b64encode(_canonical_json(certificate_payload))
    signature = server_private_key.sign(encoded_payload.encode("ascii"))
    return {"payload": encoded_payload, "signature": _b64encode(signature)}


def verify_certificate(certificate: Dict[str, str],
                       server_public_key: Ed25519PublicKey,
                       machine_private_key: Ed25519PrivateKey) -> Dict[str, Any]:
    """Verify the server signature and prove the certificate belongs to this machine."""
    if not isinstance(certificate, dict):
        raise ActivationError("Activation certificate is missing")
    encoded_payload = certificate.get("payload")
    signature = certificate.get("signature")
    if not isinstance(encoded_payload, str) or not isinstance(signature, str):
        raise ActivationError("Activation certificate is malformed")
    try:
        server_public_key.verify(_b64decode(signature), encoded_payload.encode("ascii"))
        payload = json.loads(_b64decode(encoded_payload).decode("utf-8"))
    except Exception as error:
        raise ActivationError("Activation certificate signature is invalid") from error
    if not isinstance(payload, dict) or payload.get("product") != PRODUCT_NAME:
        raise ActivationError("Activation certificate is for a different product")
    expected_machine_key = _b64encode(_public_key_bytes(machine_private_key.public_key()))
    if payload.get("machine_public_key") != expected_machine_key:
        raise ActivationError("Activation certificate belongs to a different machine")
    return payload


def _dpapi_protect(value: bytes) -> bytes:
    if os.name != "nt":
        return value
    return _dpapi_transform(value, protect=True)


def _dpapi_unprotect(value: bytes) -> bytes:
    if os.name != "nt":
        return value
    return _dpapi_transform(value, protect=False)


def _dpapi_transform(value: bytes, protect: bool) -> bytes:
    """Use Windows DPAPI so a copied device key cannot be used on another machine."""
    class DataBlob(ctypes.Structure):
        _fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_ubyte))]

    source_buffer = ctypes.create_string_buffer(value)
    source = DataBlob(len(value), ctypes.cast(source_buffer, ctypes.POINTER(ctypes.c_ubyte)))
    destination = DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    blob_pointer = ctypes.POINTER(DataBlob)
    crypt32.CryptProtectData.argtypes = [
        blob_pointer, ctypes.c_wchar_p, blob_pointer, ctypes.c_void_p,
        blob_pointer, ctypes.c_uint32, blob_pointer,
    ]
    crypt32.CryptProtectData.restype = ctypes.c_bool
    crypt32.CryptUnprotectData.argtypes = [
        blob_pointer, ctypes.POINTER(ctypes.c_wchar_p), blob_pointer, ctypes.c_void_p,
        blob_pointer, ctypes.c_uint32, blob_pointer,
    ]
    crypt32.CryptUnprotectData.restype = ctypes.c_bool
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    if protect:
        success = crypt32.CryptProtectData(
            ctypes.byref(source), None, None, None, None, 0x4, ctypes.byref(destination)
        )
    else:
        success = crypt32.CryptUnprotectData(
            ctypes.byref(source), None, None, None, None, 0, ctypes.byref(destination)
        )
    if not success:
        raise ActivationError("Windows could not protect the activation key")
    try:
        return ctypes.string_at(destination.pbData, destination.cbData)
    finally:
        kernel32.LocalFree(destination.pbData)


def default_license_directory(root: Path) -> Path:
    """Return storage outside the portable application folder."""
    if os.name == "nt":
        base = os.environ.get("PROGRAMDATA") or os.environ.get("LOCALAPPDATA")
        if base:
            return Path(base) / PRODUCT_NAME
    return Path.home() / ("." + PRODUCT_NAME.lower())


def load_or_create_device_identity(directory: Path) -> Tuple[Ed25519PrivateKey, str]:
    directory.mkdir(parents=True, exist_ok=True)
    key_path = directory / DEVICE_KEY_FILENAME
    if key_path.exists():
        try:
            private_bytes = _dpapi_unprotect(key_path.read_bytes())
            private_key = Ed25519PrivateKey.from_private_bytes(private_bytes)
            return private_key, _b64encode(_public_key_bytes(private_key.public_key()))
        except Exception as error:
            raise ActivationError("The stored machine activation key cannot be read") from error
    private_key, public_key = create_device_identity()
    private_bytes = private_key.private_bytes(
        encoding=serialization.Encoding.Raw,
        format=serialization.PrivateFormat.Raw,
        encryption_algorithm=serialization.NoEncryption(),
    )
    key_path.write_bytes(_dpapi_protect(private_bytes))
    try:
        os.chmod(str(key_path), 0o600)
    except OSError:
        pass
    return private_key, public_key


@dataclass(frozen=True)
class ActivationSettings:
    server_url: str
    public_key: str
    required: bool = True

    @property
    def enabled(self) -> bool:
        return bool(self.server_url.strip() and self.public_key.strip() and self.required)


def compiled_activation_settings() -> ActivationSettings:
    """Read build-time settings embedded in a packaged executable."""
    try:
        from src.activation_build_config import ACTIVATION_PUBLIC_KEY, ACTIVATION_SERVER
    except ImportError:
        return ActivationSettings("", "", False)
    return ActivationSettings(ACTIVATION_SERVER, ACTIVATION_PUBLIC_KEY, True)


class ActivationClient:
    """Client-side activation interface; all later checks are offline."""

    def __init__(self, server_url: str, public_key: str, state_directory: Path,
                 opener: Optional[Callable[..., Any]] = None,
                 identity: Optional[Tuple[Ed25519PrivateKey, str]] = None) -> None:
        self.server_url = server_url.rstrip("/")
        self.public_key = public_key
        self.state_directory = state_directory
        self._opener = opener or urllib.request.urlopen
        self._identity = identity

    def _device_identity(self) -> Tuple[Ed25519PrivateKey, str]:
        if self._identity is None:
            self._identity = load_or_create_device_identity(self.state_directory)
        return self._identity

    def device_public_key(self) -> str:
        """Return the public machine request value shown to the license issuer."""
        return self._device_identity()[1]

    def _server_key(self) -> Ed25519PublicKey:
        try:
            return Ed25519PublicKey.from_public_bytes(_b64decode(self.public_key))
        except Exception as error:
            raise ActivationError("The application activation key is invalid") from error

    def load_valid_certificate(self) -> Optional[Dict[str, Any]]:
        certificate_path = self.state_directory / LICENSE_FILENAME
        if not certificate_path.is_file():
            return None
        try:
            certificate = json.loads(certificate_path.read_text(encoding="utf-8"))
            private_key, _ = self._device_identity()
            return verify_certificate(certificate, self._server_key(), private_key)
        except (OSError, ValueError, ActivationError) as error:
            raise ActivationError("The saved activation is invalid or belongs to another machine") from error

    def activate(self, activation_code: str) -> Dict[str, Any]:
        if not activation_code or not activation_code.strip():
            raise ActivationError("An activation code is required")
        private_key, machine_public_key = self._device_identity()
        body = json.dumps({
            "code": activation_code.strip(),
            "machine_public_key": machine_public_key,
        }).encode("utf-8")
        request = urllib.request.Request(
            self.server_url + ACTIVATION_PATH,
            data=body,
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
                "User-Agent": "Inv_Reader-Activation/1.0",
            },
            method="POST",
        )
        try:
            with self._opener(request, timeout=30) as response:
                result = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            try:
                details = json.loads(error.read().decode("utf-8"))
                message = details.get("error", "Activation was rejected")
            except (ValueError, OSError):
                message = "Activation was rejected"
            raise ActivationError(message) from error
        except (urllib.error.URLError, OSError, ValueError) as error:
            raise ActivationError("Could not connect to the activation server") from error
        certificate = result.get("certificate") if isinstance(result, dict) else None
        payload = verify_certificate(certificate, self._server_key(), private_key)
        self.state_directory.mkdir(parents=True, exist_ok=True)
        (self.state_directory / LICENSE_FILENAME).write_text(
            json.dumps(certificate, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
        )
        return payload
