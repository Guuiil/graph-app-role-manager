# Permissions and safety

Graph App Role Manager changes application permissions assigned to identities. Treat it as
an administrative tool.

## Delegated Microsoft Graph permissions

| Permission | Purpose |
| --- | --- |
| `Application.Read.All` | Search service principals, load Graph app roles, and inspect assignments |
| `AppRoleAssignment.ReadWrite.All` | Create and remove app-role assignments |

The signed-in administrator must also hold an appropriate Microsoft Entra role permitted to
manage the target service principal. Organizational policies and protected applications can
impose additional restrictions.

## Before changing an assignment

- Confirm the tenant used for authentication.
- Verify the target's display name, object ID, application ID, and service-principal type.
- Check the workload's documented permission requirements.
- Prefer the least-privileged application permission.
- Understand that application permissions are tenant-wide unless the target service applies
  an additional resource-level authorization model.

For example, assigning `Sites.Selected` alone does not grant access to a SharePoint site; a
separate site-level grant is still required.

## Removal

Removing an app-role assignment can immediately break automation or applications. The tool
therefore requires the operator to select existing Microsoft Graph assignments and approve
a confirmation dialog. It does not remove permissions from other resource APIs.

## Credential handling

- The Windows version uses the Microsoft Graph PowerShell authentication context for the
  current process.
- The cross-platform Python interface uses a process-scoped Microsoft Graph PowerShell
  authentication context in its bundled backend.
- Neither version asks for or stores a password, client secret, or certificate.

Do not publish screenshots or logs containing tenant identifiers, account names, object IDs,
or other organizational information without reviewing them first.
