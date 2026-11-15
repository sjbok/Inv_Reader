import base64
import json
import tempfile
import unittest
from pathlib import Path

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives import serialization

from activation_server.service import ActivationService
from src.licensing import (
    ActivationClient,
    ActivationError,
    create_certificate,
    create_device_identity,
    load_or_create_device_identity,
    verify_certificate,
)


class LicensingTests(unittest.TestCase):
    def test_certificate_can_be_verified_for_the_activated_machine(self):
        server_private_key = Ed25519PrivateKey.generate()
        machine_private_key, machine_public_key = create_device_identity()
        certificate = create_certificate(
            server_private_key,
            {
                "client_name": "Example Company",
                "license_id": "example-company",
                "machine_public_key": machine_public_key,
                "product": "Inv_Reader",
            },
        )

        verified = verify_certificate(
            certificate,
            server_private_key.public_key(),
            machine_private_key,
        )

        self.assertEqual(verified["client_name"], "Example Company")
        self.assertEqual(verified["machine_public_key"], machine_public_key)

    def test_tampered_certificate_is_rejected(self):
        server_private_key = Ed25519PrivateKey.generate()
        machine_private_key, machine_public_key = create_device_identity()
        certificate = create_certificate(
            server_private_key,
            {"machine_public_key": machine_public_key, "product": "Inv_Reader"},
        )
        payload = json.loads(base64.urlsafe_b64decode(certificate["payload"] + "=="))
        payload["product"] = "DifferentProduct"
        certificate["payload"] = base64.urlsafe_b64encode(
            json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
        ).decode("ascii").rstrip("=")

        with self.assertRaises(ActivationError):
            verify_certificate(certificate, server_private_key.public_key(), machine_private_key)

    def test_certificate_cannot_be_used_on_another_machine(self):
        server_private_key = Ed25519PrivateKey.generate()
        first_private_key, first_public_key = create_device_identity()
        second_private_key, _ = create_device_identity()
        certificate = create_certificate(
            server_private_key,
            {"machine_public_key": first_public_key, "product": "Inv_Reader"},
        )

        with self.assertRaises(ActivationError):
            verify_certificate(certificate, server_private_key.public_key(), second_private_key)

    def test_activation_code_is_consumed_for_one_machine_only(self):
        with tempfile.TemporaryDirectory() as directory:
            service = ActivationService(Path(directory))
            _, machine_public_key = create_device_identity()
            _, other_machine_public_key = create_device_identity()
            issued = service.create_activation_code(
                "Example Company", "example-company", machine_public_key
            )

            with self.assertRaises(ActivationError):
                service.activate(issued["code"], other_machine_public_key)

            result = service.activate(issued["code"], machine_public_key)
            payload = json.loads(base64.urlsafe_b64decode(
                result["certificate"]["payload"] + "=="
            ).decode("utf-8"))
            self.assertEqual(payload["client_name"], "Example Company")
            retry = service.activate(issued["code"], machine_public_key)
            self.assertEqual(retry, result)
            with self.assertRaises(ActivationError):
                service.activate(issued["code"], create_device_identity()[1])

    def test_client_stores_certificate_after_activation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            server_private_key = Ed25519PrivateKey.generate()
            machine_private_key, machine_public_key = load_or_create_device_identity(root)
            certificate = create_certificate(
                server_private_key,
                {"machine_public_key": machine_public_key, "product": "Inv_Reader"},
            )
            client = ActivationClient(
                "https://activation.example.test",
                base64.urlsafe_b64encode(
                    server_private_key.public_key().public_bytes(
                        serialization.Encoding.Raw, serialization.PublicFormat.Raw
                    )
                ).decode("ascii").rstrip("="),
                root,
                opener=lambda request, timeout: type(
                    "Response",
                    (),
                    {
                        "__enter__": lambda self: self,
                        "__exit__": lambda self, *args: False,
                        "read": lambda self: json.dumps({"certificate": certificate}).encode("utf-8"),
                    },
                )(),
                identity=(machine_private_key, machine_public_key),
            )

            client.activate("INV-test-code")

            self.assertTrue((root / "license.json").is_file())

            offline_client = ActivationClient(
                "https://server-is-not-needed.example.test",
                client.public_key,
                root,
            )
            self.assertEqual(
                offline_client.load_valid_certificate()["product"],
                "Inv_Reader",
            )


if __name__ == "__main__":
    unittest.main()
