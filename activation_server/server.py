"""HTTP adapter for the Inv_Reader activation service.

Run locally with:
    ACTIVATION_ADMIN_TOKEN=change-me python -m activation_server.server
"""

import argparse
import json
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, Optional

from src.licensing import ActivationError
from .service import ActivationService


class ActivationRequestHandler(BaseHTTPRequestHandler):
    service: ActivationService
    admin_token: str

    def log_message(self, format: str, *args: Any) -> None:
        print("activation-server: " + format % args)

    def _send_json(self, status: int, value: Dict[str, Any]) -> None:
        body = json.dumps(value, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> Dict[str, Any]:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            value = json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, TypeError, json.JSONDecodeError) as error:
            raise ActivationError("Request body must be valid JSON") from error
        if not isinstance(value, dict):
            raise ActivationError("Request body must be a JSON object")
        return value

    def _authorized(self) -> bool:
        expected = "Bearer " + self.admin_token
        return self.headers.get("Authorization", "") == expected

    def do_GET(self) -> None:
        if self.path == "/health":
            self._send_json(200, {"status": "ok"})
            return
        if self.path == "/v1/public-key":
            self._send_json(200, {"public_key": self.service.public_key()})
            return
        if self.path == "/v1/admin/codes":
            if not self._authorized():
                self._send_json(401, {"error": "Admin authorization required"})
                return
            self._send_json(200, {"codes": self.service.list_activation_codes()})
            return
        self._send_json(404, {"error": "Not found"})

    def do_POST(self) -> None:
        try:
            body = self._read_json()
            if self.path == "/v1/activate":
                self._send_json(200, self.service.activate(
                    str(body.get("code", "")), str(body.get("machine_public_key", ""))
                ))
                return
            if self.path == "/v1/admin/codes":
                if not self._authorized():
                    self._send_json(401, {"error": "Admin authorization required"})
                    return
                self._send_json(201, {"activation": self.service.create_activation_code(
                    str(body.get("client_name", "")), body.get("license_id"),
                    body.get("machine_public_key")
                )})
                return
            self._send_json(404, {"error": "Not found"})
        except ActivationError as error:
            self._send_json(400, {"error": str(error)})
        except Exception as error:
            print("activation-server error:", error)
            self._send_json(500, {"error": "The activation server encountered an error"})


def create_server(service: ActivationService, host: str, port: int,
                  admin_token: str) -> ThreadingHTTPServer:
    if not admin_token:
        raise RuntimeError("ACTIVATION_ADMIN_TOKEN must be set")
    handler = type("ConfiguredActivationRequestHandler", (ActivationRequestHandler,), {
        "service": service,
        "admin_token": admin_token,
    })
    return ThreadingHTTPServer((host, port), handler)


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="Run the Inv_Reader activation server")
    parser.add_argument("--host", default=os.environ.get("HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.environ.get("PORT", "8765")))
    parser.add_argument("--data", type=Path,
                        default=Path(os.environ.get("ACTIVATION_DATA_DIR", "activation_server_data")))
    parser.add_argument("--admin-token", default=os.environ.get("ACTIVATION_ADMIN_TOKEN"))
    args = parser.parse_args(argv)
    service = ActivationService(args.data)
    server = create_server(service, args.host, args.port, args.admin_token or "")
    print("Activation server listening on http://{}:{}".format(args.host, args.port))
    print("Activation public key: {}".format(service.public_key()))
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("Stopping activation server")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
