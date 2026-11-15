"""Persistent activation-code and certificate service."""

import hashlib
from datetime import datetime
import json
from pathlib import Path
import secrets
import sqlite3
from typing import Any, Dict, List, Optional
import uuid

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from src.licensing import ActivationError, create_certificate, _b64encode, _b64decode


class ActivationService:
    """Business interface used by both the HTTP server and tests."""

    def __init__(self, data_directory: Path) -> None:
        self.data_directory = data_directory
        self.data_directory.mkdir(parents=True, exist_ok=True)
        self.database_path = self.data_directory / "activations.sqlite3"
        self.key_path = self.data_directory / "signing-key.bin"
        self._private_key = self._load_or_create_signing_key()
        self._create_tables()

    def _load_or_create_signing_key(self) -> Ed25519PrivateKey:
        if self.key_path.exists():
            return Ed25519PrivateKey.from_private_bytes(self.key_path.read_bytes())
        private_key = Ed25519PrivateKey.generate()
        self.key_path.write_bytes(private_key.private_bytes(
            encoding=serialization.Encoding.Raw,
            format=serialization.PrivateFormat.Raw,
            encryption_algorithm=serialization.NoEncryption(),
        ))
        try:
            self.key_path.chmod(0o600)
        except OSError:
            pass
        return private_key

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(str(self.database_path))
        connection.row_factory = sqlite3.Row
        return connection

    def _create_tables(self) -> None:
        with self._connect() as connection:
            connection.execute("""
                CREATE TABLE IF NOT EXISTS activation_codes (
                    activation_id TEXT PRIMARY KEY,
                    code_hash TEXT NOT NULL UNIQUE,
                    client_name TEXT NOT NULL,
                    license_id TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    used_at TEXT,
                    machine_public_key TEXT,
                    certificate_json TEXT,
                    requested_machine_public_key TEXT
                )
            """)
            columns = {
                row[1] for row in connection.execute("PRAGMA table_info(activation_codes)")
            }
            if "certificate_json" not in columns:
                connection.execute("ALTER TABLE activation_codes ADD COLUMN certificate_json TEXT")
            if "requested_machine_public_key" not in columns:
                connection.execute(
                    "ALTER TABLE activation_codes ADD COLUMN requested_machine_public_key TEXT"
                )

    @staticmethod
    def _hash_code(code: str) -> str:
        return hashlib.sha256(code.encode("utf-8")).hexdigest()

    def create_activation_code(self, client_name: str, license_id: Optional[str] = None,
                               requested_machine_public_key: Optional[str] = None) -> Dict[str, str]:
        if not client_name.strip():
            raise ActivationError("Client name is required")
        if requested_machine_public_key:
            self._validate_machine_public_key(requested_machine_public_key)
        license_value = str(license_id).strip() if license_id else client_name.strip()
        code = "INV-" + secrets.token_urlsafe(18)
        activation_id = str(uuid.uuid4())
        created_at = datetime.utcnow().isoformat() + "Z"
        with self._connect() as connection:
            connection.execute(
                "INSERT INTO activation_codes VALUES (?, ?, ?, ?, ?, NULL, NULL, NULL, ?)",
                (activation_id, self._hash_code(code), client_name.strip(),
                 license_value, created_at, requested_machine_public_key),
            )
        return {"activation_id": activation_id, "code": code, "client_name": client_name.strip()}

    def activate(self, code: str, machine_public_key: str) -> Dict[str, Any]:
        if not code or not machine_public_key:
            raise ActivationError("Activation code and machine identity are required")
        self._validate_machine_public_key(machine_public_key)
        now = datetime.utcnow().isoformat() + "Z"
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT * FROM activation_codes WHERE code_hash = ?", (self._hash_code(code.strip()),)
            ).fetchone()
            if row is None:
                raise ActivationError("Activation code is invalid")
            if (
                row["requested_machine_public_key"]
                and row["requested_machine_public_key"] != machine_public_key
            ):
                raise ActivationError("Activation code is for a different machine")
            if row["used_at"] is not None:
                if row["machine_public_key"] == machine_public_key and row["certificate_json"]:
                    return {"certificate": json.loads(row["certificate_json"])}
                raise ActivationError("Activation code has already been used")
            certificate = create_certificate(self._private_key, {
                "activation_id": row["activation_id"],
                "client_name": row["client_name"],
                "license_id": row["license_id"],
                "machine_public_key": machine_public_key,
            })
            connection.execute(
                "UPDATE activation_codes SET used_at = ?, machine_public_key = ?, certificate_json = ? "
                "WHERE activation_id = ?",
                (now, machine_public_key, json.dumps(certificate), row["activation_id"]),
            )
        return {"certificate": certificate}

    def public_key(self) -> str:
        return _b64encode(self._private_key.public_key().public_bytes(
            encoding=serialization.Encoding.Raw,
            format=serialization.PublicFormat.Raw,
        ))

    def list_activation_codes(self) -> List[Dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT activation_id, client_name, license_id, created_at, used_at, machine_public_key, "
                "requested_machine_public_key "
                "FROM activation_codes ORDER BY created_at"
            ).fetchall()
        return [dict(row) for row in rows]

    @staticmethod
    def _validate_machine_public_key(machine_public_key: str) -> None:
        try:
            from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
            Ed25519PublicKey.from_public_bytes(_b64decode(machine_public_key))
        except Exception as error:
            raise ActivationError("Machine identity is invalid") from error
