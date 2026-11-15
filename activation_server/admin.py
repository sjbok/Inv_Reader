"""Command-line administration for a deployed activation server."""

import argparse
import json
import os
from typing import Any, Dict, Optional
import urllib.request


def _request(base_url: str, path: str, token: Optional[str] = None,
             body: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    data = json.dumps(body).encode("utf-8") if body is not None else None
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = "Bearer " + token
    request = urllib.request.Request(base_url.rstrip("/") + path, data=data, headers=headers,
                                     method="POST" if body is not None else "GET")
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="Administer Inv_Reader machine activations")
    parser.add_argument("--server", default=os.environ.get("ACTIVATION_SERVER_URL"), required=False)
    parser.add_argument("--admin-token", default=os.environ.get("ACTIVATION_ADMIN_TOKEN"))
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("public-key")
    create = subparsers.add_parser("create-code")
    create.add_argument("--client", required=True, dest="client_name")
    create.add_argument("--machine-public-key", required=True)
    create.add_argument("--license-id", default=None)
    subparsers.add_parser("list")
    args = parser.parse_args(argv)
    if not args.server:
        parser.error("--server or ACTIVATION_SERVER_URL is required")
    if args.command != "public-key" and not args.admin_token:
        parser.error("--admin-token or ACTIVATION_ADMIN_TOKEN is required")

    if args.command == "public-key":
        result = _request(args.server, "/v1/public-key")
        print(result["public_key"])
    elif args.command == "create-code":
        result = _request(args.server, "/v1/admin/codes", args.admin_token, {
            "client_name": args.client_name,
            "license_id": args.license_id,
            "machine_public_key": args.machine_public_key,
        })
        print("Give this one-time code to the client: {}".format(result["activation"]["code"]))
        print("Activation ID: {}".format(result["activation"]["activation_id"]))
    else:
        result = _request(args.server, "/v1/admin/codes", args.admin_token)
        print(json.dumps(result["codes"], indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
