# Cross-platform setup

## Requirements

- Python 3.10 or later.
- A working Tkinter installation (`python -m tkinter` should open a test window).
- The packages in `cross-platform/requirements.txt`.
- A tenant-owned public-client app registration.

## Create the public-client app registration

Create an App Registration dedicated to this administrative tool:

1. Use a single-tenant supported account type unless the organization requires otherwise.
2. Enable public-client flows.
3. Add the delegated Microsoft Graph permissions:
   - `Application.Read.All`
   - `AppRoleAssignment.ReadWrite.All`
4. Grant administrator consent.
5. Record the Application (client) ID and tenant ID.

Do not create a client secret. Device-code flow is intended for public clients that cannot
protect a secret.

## Install

macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r cross-platform/requirements.txt
python cross-platform/graph_app_role_manager.py
```

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r .\cross-platform\requirements.txt
python .\cross-platform\graph_app_role_manager.py
```

## Authenticate

Enter:

- the tenant ID or verified tenant domain;
- the Application (client) ID of the public-client app registration.

The interface starts device-code authentication, copies the short-lived user code to the
clipboard, and opens Microsoft's verification page. The resulting access token remains in
memory and is discarded when the process exits.

## Troubleshooting

- `AADSTS65001`: administrator consent has not been granted.
- HTTP `403`: the session lacks a required Graph permission, the signed-in user lacks an
  appropriate Entra role, or the target assignment is protected.
- No Tkinter window: install the Tk support supplied for the operating system's Python
  distribution.
- No search results: use the beginning of the target service principal's display name.
