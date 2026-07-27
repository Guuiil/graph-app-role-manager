# Windows setup

## Requirements

- Windows 10/11 or Windows Server with a desktop session.
- PowerShell 7 or later.
- `Microsoft.Graph.Authentication`.
- `Microsoft.Graph.Applications`.
- An administrator account authorized to manage app-role assignments.

## Start

From the repository root:

```powershell
pwsh -File .\windows\Graph-App-Role-Manager.ps1
```

If a required Microsoft Graph module is missing, the interface displays the exact module
names and asks whether it may install them for the current user. Declining leaves the
machine unchanged and stops the connection.

Modules can also be installed beforehand:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Applications -Scope CurrentUser
```

## Authentication

The tool calls `Connect-MgGraph` with:

- `Application.Read.All`
- `AppRoleAssignment.ReadWrite.All`

Authentication is process-scoped. The Graph session is disconnected when the form closes.

## Usage

1. Select **Connect to Microsoft Graph**.
2. Enter part of the managed identity or service-principal display name.
3. Select the correct result and verify its identifiers.
4. Filter the Graph permission catalog.
5. Check permissions and select **Assign selected permissions**.
6. To remove assignments, select rows in the current-assignment table and choose
   **Remove selected**.

Every change displays the target and permissions in a confirmation dialog.
