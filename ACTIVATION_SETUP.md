# Inv_Reader activation setup

This project includes a small activation server. It issues one-time codes for
individual client machines. The Windows application contacts the server only
for its first activation; afterward it verifies the signed license offline.

## Important security rules

- Keep `activation_server_data/signing-key.bin` private and back it up securely.
- Never commit `activation_server_data/` to Git.
- Never put the signing key in the Windows application.
- Use an HTTPS URL when the server is reachable over the internet.
- Create one activation code per approved machine.
- Treat the activation code as a secret and give it only to the client.

## Test locally first

From the repository root, install the requirements and choose an admin token:

```bash
python3 -m venv .activation-venv
source .activation-venv/bin/activate
pip install -r requirements.txt
export ACTIVATION_ADMIN_TOKEN="choose-a-long-random-secret"
python -m activation_server.server
```

The server prints an activation public key. Leave this terminal running.

In another terminal, create one machine code:

```bash
export ACTIVATION_SERVER_URL="http://127.0.0.1:8765"
export ACTIVATION_ADMIN_TOKEN="choose-a-long-random-secret"
python -m activation_server.admin create-code \
  --client "Example Company" \
  --machine-public-key "MACHINE-KEY-SHOWN-BY-THE-APP"
```

The client application shows the machine key before asking for the activation
code. The client must send that key to you first. The command then prints a
one-time code bound to that exact machine. Do not reuse that code for a second
machine. Create another code, with a new machine key, for every additional
approved machine.

You can inspect the public key with:

```bash
python -m activation_server.admin public-key
```

## Deploy the server

Deploy the `activation_server` package, the `src` package, and
`requirements.txt` to a small HTTPS Python host. The host must provide
persistent storage for:

```text
activation_server_data/activations.sqlite3
activation_server_data/signing-key.bin
```

Set these environment variables on the host:

```text
ACTIVATION_ADMIN_TOKEN=<long random secret>
ACTIVATION_DATA_DIR=/persistent/path/activation_server_data
```

Run the web process with:

```bash
python -m activation_server.server
```

Use the host's HTTPS URL for the application. Do not expose the development
server over plain HTTP on the public internet. If the hosting provider uses a
temporary filesystem, add a persistent disk or the database and signing key
will be lost during a restart.

After deployment, retrieve the public key:

```bash
ACTIVATION_SERVER_URL="https://your-server.example" \
ACTIVATION_ADMIN_TOKEN="your-admin-token" \
python -m activation_server.admin public-key
```

## Build the client application

On the Windows build computer, pass both the HTTPS server URL and public key:

```powershell
.\build_windows.ps1 `
  -ActivationServer "https://your-server.example" `
  -ActivationPublicKey "PASTE_PUBLIC_KEY_HERE"
```

The resulting package requires activation before it processes documents.

## Client activation workflow

1. The client starts the new EXE and sends you the displayed machine key.
2. Create one code with `create-code`, including that machine key.
3. Give the code to the client for that approved computer.
4. The client enters the code.
5. The server consumes the code and binds it to that machine.
6. The client can then use the application offline.
7. For another computer, repeat the process and charge for another code.

If a machine is replaced, the current basic server does not automatically
free the old slot. You should record the old activation ID, then either issue
a replacement code after confirming the old machine is no longer used, or add
a reset/revocation workflow before production use.

## Old test copies

The existing test copy has no activation check. Replace it with a newly built
package and tell the client in writing that the old copy is an evaluation copy
that may not be redistributed and must be deleted after replacement. A local
activation system cannot retroactively disable an old build that never checked
with the server.
