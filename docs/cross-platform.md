# Cross-platform v1.0 setup

## Requirements

- Windows, macOS, or a Linux desktop environment.
- Python 3.10 or later.
- Tkinter (`python -m tkinter` should open a test window).
- PowerShell 7 (`pwsh`).
- The `Microsoft.Graph.Authentication` PowerShell module.

No custom App Registration, client ID, client secret, certificate, MSAL package, or
`requests` package is required.

## Install the Graph module

Run in PowerShell 7:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

If your organization manages PowerShell repositories centrally, follow its approved module
installation process instead.

## Start the application

From the repository root:

```bash
python cross-platform/graph_app_role_manager.py
```

Keep these files in the same directory:

- `graph_app_role_manager.py`
- `graph_backend.ps1`
- `graph-app-role-manager-icon.png`

The Python standard library is sufficient, so creating a virtual environment is optional.

## Authenticate

Select **Connect to Microsoft Graph**. The bundled PowerShell backend starts an interactive
`Connect-MgGraph` sign-in and requests:

- `Application.Read.All`
- `AppRoleAssignment.ReadWrite.All`

Leave the tenant field empty for the normal organizational sign-in flow, or enter a tenant
ID or verified tenant domain to require a specific tenant. The authentication context is
limited to the backend process and is disconnected when the interface closes.

## Use the v1.0 interface

1. Connect with an authorized administrator account.
2. Select an identity from the list loaded after sign-in, type to filter it locally, or use
   **Search** for a server-side display-name search.
3. Verify the target name, type, status, object ID, and application ID.
4. Filter and select Microsoft Graph application permissions. Selections remain in the
   basket while the filter changes.
5. Review the in-app confirmation and assign the permissions.
6. Select existing assignments and confirm removal when needed.

## Troubleshooting

Check PowerShell and the Graph module:

```powershell
$PSVersionTable.PSVersion
Get-Module -ListAvailable Microsoft.Graph.Authentication
```

Test authentication separately:

```powershell
Connect-MgGraph -Scopes Application.Read.All,AppRoleAssignment.ReadWrite.All -NoWelcome
Get-MgContext | Format-List
Disconnect-MgGraph
```

- If `pwsh` is not found, install PowerShell 7 and restart the application.
- If the browser does not open, run the independent authentication test above and inspect
  its full error.
- `AADSTS65001` usually means administrator consent has not been granted.
- HTTP `403` indicates missing Graph permission, an insufficient Entra administrator role,
  an organizational policy restriction, or a protected target.
- If no Tkinter window opens, install the Tk support supplied for your Python distribution.
