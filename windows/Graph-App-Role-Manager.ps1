#Requires -Version 7.0
<#
.SYNOPSIS
Graph App Role Manager - GUI helper for assigning Microsoft Graph application
permissions (app roles) to a Managed Identity or another service principal.

.DESCRIPTION
Windows Forms GUI that:
- Connects to Microsoft Graph with delegated administrative scopes.
- Searches a target service principal by display name.
- Lists Microsoft Graph application permissions.
- Shows permissions already assigned to the target.
- Assigns one or more selected application permissions.
- Removes one or more selected application permissions after confirmation.
- Avoids duplicate assignments and writes an execution log.

REQUIREMENTS
- Windows
- PowerShell 7+
- Microsoft.Graph.Authentication
- Microsoft.Graph.Applications

The signed-in administrator must be allowed to create app role assignments.
Typical delegated scopes requested by this tool:
- Application.Read.All
- AppRoleAssignment.ReadWrite.All

SECURITY
Grant only the permissions needed by the target workload.
#>

param(
    [switch]$NoGui
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $IsWindows) {
    throw 'This GUI uses Windows Forms and must be run on Windows.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# STATE
# ---------------------------------------------------------------------------

$script:GraphConnected         = $false
$script:GraphServicePrincipal  = $null
$script:TargetServicePrincipal = $null
$script:AllApplicationRoles    = @()
$script:VisibleApplicationRoles = @()

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

function Write-UiLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $line = "[$timestamp][$Level] $Message"

    if ($null -ne $txtLog) {
        $txtLog.AppendText($line + [Environment]::NewLine)
        $txtLog.SelectionStart = $txtLog.TextLength
        $txtLog.ScrollToCaret()
    }

    Write-Host $line
}

function Show-UiError {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-UiLog -Message $Message -Level 'ERROR'
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        'Graph App Role Manager',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Set-BusyState {
    param([bool]$Busy)

    $form.UseWaitCursor = $Busy

    foreach ($control in @(
        $btnConnect,
        $btnSearchIdentity,
        $btnLoadPermissions,
        $btnAssign,
        $btnRemove,
        $btnRefreshAssignments
    )) {
        if ($null -ne $control) {
            $control.Enabled = -not $Busy
        }
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Ensure-GraphModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Applications'
    )

    # Always preserve collection semantics. Without @(...), PowerShell unwraps a
    # single missing module to a scalar string and StrictMode rejects .Count.
    $missingModules = @(
        foreach ($moduleName in $requiredModules) {
            if (-not (Get-Module -ListAvailable -Name $moduleName)) {
                $moduleName
            }
        }
    )

    if ($missingModules.Count -gt 0) {
        $message = @"
The following PowerShell modules are missing:

$($missingModules -join [Environment]::NewLine)

Install them for the current user now?
"@

        $answer = [System.Windows.Forms.MessageBox]::Show(
            $message,
            'Missing Microsoft Graph modules',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            throw 'Required Microsoft Graph modules are not installed.'
        }

        foreach ($moduleName in $missingModules) {
            Write-UiLog -Message "Installing module: $moduleName"
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
        }
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
}

function Escape-ODataString {
    param([Parameter(Mandatory)][string]$Value)
    return $Value.Replace("'", "''")
}

function Get-SelectedTargetServicePrincipal {
    if ($null -eq $cmbIdentityResults.SelectedItem) {
        return $null
    }

    return $cmbIdentityResults.SelectedItem.Tag
}

function Update-IdentitySummary {
    param($ServicePrincipal)

    if ($null -eq $ServicePrincipal) {
        $lblIdentityNameValue.Text = '-'
        $lblObjectIdValue.Text = '-'
        $lblAppIdValue.Text = '-'
        return
    }

    $lblIdentityNameValue.Text = [string]$ServicePrincipal.DisplayName
    $lblObjectIdValue.Text = [string]$ServicePrincipal.Id
    $lblAppIdValue.Text = [string]$ServicePrincipal.AppId
}

function Load-GraphApplicationRoles {
    if (-not $script:GraphConnected) {
        throw 'Connect to Microsoft Graph first.'
    }

    Write-UiLog -Message 'Loading Microsoft Graph application permissions...'

    $script:GraphServicePrincipal = Get-MgServicePrincipal `
        -Filter "appId eq '00000003-0000-0000-c000-000000000000'" `
        -Property 'id,appId,displayName,appRoles' `
        -ConsistencyLevel eventual `
        -ErrorAction Stop |
        Select-Object -First 1

    if ($null -eq $script:GraphServicePrincipal) {
        throw 'The Microsoft Graph service principal could not be found.'
    }

    $script:AllApplicationRoles = @(
        $script:GraphServicePrincipal.AppRoles |
        Where-Object {
            $_.IsEnabled -eq $true -and
            $_.AllowedMemberTypes -contains 'Application' -and
            -not [string]::IsNullOrWhiteSpace($_.Value)
        } |
        Sort-Object Value
    )

    Apply-PermissionFilter

    Write-UiLog -Message (
        "Loaded {0} Microsoft Graph application permissions." -f
        $script:AllApplicationRoles.Count
    ) -Level 'SUCCESS'
}

function Apply-PermissionFilter {
    $filter = $txtPermissionFilter.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($filter)) {
        $script:VisibleApplicationRoles = @($script:AllApplicationRoles)
    }
    else {
        $script:VisibleApplicationRoles = @(
            $script:AllApplicationRoles |
            Where-Object {
                $_.Value -like "*$filter*" -or
                $_.DisplayName -like "*$filter*" -or
                $_.Description -like "*$filter*"
            }
        )
    }

    $checkedValues = @{}
    foreach ($index in $checkedPermissions.CheckedIndices) {
        $role = $checkedPermissions.Items[$index]
        $checkedValues[[string]$role.Value] = $true
    }

    $checkedPermissions.BeginUpdate()
    try {
        $checkedPermissions.Items.Clear()

        foreach ($role in $script:VisibleApplicationRoles) {
            $item = [pscustomobject]@{
                Value       = [string]$role.Value
                DisplayName = [string]$role.DisplayName
                Description = [string]$role.Description
                Id          = [guid]$role.Id
                Text        = "{0} — {1}" -f $role.Value, $role.DisplayName
            }

            $newIndex = $checkedPermissions.Items.Add($item)
            if ($checkedValues.ContainsKey($item.Value)) {
                $checkedPermissions.SetItemChecked($newIndex, $true)
            }
        }
    }
    finally {
        $checkedPermissions.EndUpdate()
    }

    $lblPermissionCount.Text = "{0} permission(s) displayed" -f $script:VisibleApplicationRoles.Count
}

function Refresh-CurrentAssignments {
    if ($null -eq $script:TargetServicePrincipal) {
        throw 'Select a target Managed Identity or service principal first.'
    }

    if ($null -eq $script:GraphServicePrincipal) {
        Load-GraphApplicationRoles
    }

    Write-UiLog -Message "Reading existing assignments for '$($script:TargetServicePrincipal.DisplayName)'..."

    $assignments = @(
        Get-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $script:TargetServicePrincipal.Id `
            -All `
            -ErrorAction Stop |
        Where-Object {
            $_.ResourceId -eq $script:GraphServicePrincipal.Id
        }
    )

    $roleById = @{}
    foreach ($role in $script:AllApplicationRoles) {
        $roleById[[string]$role.Id] = $role
    }

    $gridAssignments.Rows.Clear()

    foreach ($assignment in $assignments | Sort-Object AppRoleId) {
        $role = $roleById[[string]$assignment.AppRoleId]

        $permissionValue = if ($null -ne $role) {
            [string]$role.Value
        }
        else {
            "Unknown role: $($assignment.AppRoleId)"
        }

        $displayName = if ($null -ne $role) {
            [string]$role.DisplayName
        }
        else {
            ''
        }

        [void]$gridAssignments.Rows.Add(
            $permissionValue,
            $displayName,
            [string]$assignment.AppRoleId,
            [string]$assignment.Id
        )
    }

    $lblAssignedCount.Text = "{0} Microsoft Graph permission(s) assigned" -f $assignments.Count
    Write-UiLog -Message "Current assignments loaded: $($assignments.Count)." -Level 'SUCCESS'
}

function Search-TargetServicePrincipals {
    $name = $txtIdentityName.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($name)) {
        throw 'Enter a Managed Identity or service principal display name.'
    }

    if (-not $script:GraphConnected) {
        throw 'Connect to Microsoft Graph first.'
    }

    $escaped = Escape-ODataString -Value $name

    Write-UiLog -Message "Searching service principals matching: $name"

    # StartsWith supports partial searches and is convenient for Managed Identity names.
    $results = @(
        Get-MgServicePrincipal `
            -Filter "startsWith(displayName,'$escaped')" `
            -Property 'id,appId,displayName,servicePrincipalType,accountEnabled' `
            -ConsistencyLevel eventual `
            -All `
            -ErrorAction Stop |
        Sort-Object DisplayName
    )

    $cmbIdentityResults.Items.Clear()
    $script:TargetServicePrincipal = $null
    Update-IdentitySummary -ServicePrincipal $null
    $gridAssignments.Rows.Clear()

    foreach ($sp in $results) {
        $label = "{0} | {1} | {2}" -f (
            $sp.DisplayName,
            $sp.ServicePrincipalType,
            $sp.Id
        )

        # ComboBox cannot directly host controls; use a PSObject with ToString().
        $item = [pscustomobject]@{
            Text = $label
            Tag  = $sp
        }
        $item | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Text } -Force

        [void]$cmbIdentityResults.Items.Add($item)
    }

    if ($results.Count -eq 0) {
        Write-UiLog -Message 'No matching service principal was found.' -Level 'WARNING'
        return
    }

    $cmbIdentityResults.SelectedIndex = 0
    Write-UiLog -Message "Found $($results.Count) matching service principal(s)." -Level 'SUCCESS'
}

function Assign-SelectedPermissions {
    if ($null -eq $script:TargetServicePrincipal) {
        throw 'Select a target Managed Identity or service principal first.'
    }

    if ($null -eq $script:GraphServicePrincipal) {
        Load-GraphApplicationRoles
    }

    $selectedItems = @()
    foreach ($item in $checkedPermissions.CheckedItems) {
        $selectedItems += $item
    }

    if ($selectedItems.Count -eq 0) {
        throw 'Select at least one Microsoft Graph application permission.'
    }

    $confirmationMessage = @"
Target:
$($script:TargetServicePrincipal.DisplayName)

Permissions to process:
$($selectedItems.Value -join [Environment]::NewLine)

Proceed with the app role assignments?
"@

    $answer = [System.Windows.Forms.MessageBox]::Show(
        $confirmationMessage,
        'Confirm permission assignment',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-UiLog -Message 'Permission assignment cancelled by the user.' -Level 'WARNING'
        return
    }

    $existingAssignments = @(
        Get-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $script:TargetServicePrincipal.Id `
            -All `
            -ErrorAction Stop |
        Where-Object {
            $_.ResourceId -eq $script:GraphServicePrincipal.Id
        }
    )

    $existingRoleIds = @{}
    foreach ($assignment in $existingAssignments) {
        $existingRoleIds[[string]$assignment.AppRoleId] = $true
    }

    $assigned = 0
    $skipped  = 0
    $failed   = 0

    foreach ($item in $selectedItems) {
        $roleId = [string]$item.Id

        if ($existingRoleIds.ContainsKey($roleId)) {
            Write-UiLog -Message "Already assigned: $($item.Value)" -Level 'WARNING'
            $skipped++
            continue
        }

        try {
            $body = @{
                PrincipalId = [guid]$script:TargetServicePrincipal.Id
                ResourceId  = [guid]$script:GraphServicePrincipal.Id
                AppRoleId   = [guid]$item.Id
            }

            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $script:TargetServicePrincipal.Id `
                -BodyParameter $body `
                -ErrorAction Stop | Out-Null

            Write-UiLog -Message "Assigned successfully: $($item.Value)" -Level 'SUCCESS'
            $existingRoleIds[$roleId] = $true
            $assigned++
        }
        catch {
            Write-UiLog -Message (
                "Failed to assign {0}: {1}" -f
                $item.Value,
                $_.Exception.Message
            ) -Level 'ERROR'
            $failed++
        }
    }

    Refresh-CurrentAssignments

    [System.Windows.Forms.MessageBox]::Show(
        "Completed.`n`nAssigned: $assigned`nAlready present: $skipped`nFailed: $failed",
        'Graph App Role Manager',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Remove-SelectedPermissions {
    if ($null -eq $script:TargetServicePrincipal) {
        throw 'Select a target Managed Identity or service principal first.'
    }

    $selectedRows = @($gridAssignments.SelectedRows)
    if ($selectedRows.Count -eq 0) {
        throw 'Select at least one existing Microsoft Graph permission to remove.'
    }

    $selectedAssignments = @(
        foreach ($row in $selectedRows) {
            [pscustomobject]@{
                Permission   = [string]$row.Cells['Permission'].Value
                AssignmentId = [string]$row.Cells['AssignmentId'].Value
            }
        }
    )

    $confirmationMessage = @"
Target:
$($script:TargetServicePrincipal.DisplayName)

Permissions to remove:
$($selectedAssignments.Permission -join [Environment]::NewLine)

Remove these app role assignments?
"@

    $answer = [System.Windows.Forms.MessageBox]::Show(
        $confirmationMessage,
        'Confirm permission removal',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-UiLog -Message 'Permission removal cancelled by the user.' -Level 'WARNING'
        return
    }

    $removed = 0
    $failed  = 0

    foreach ($assignment in $selectedAssignments) {
        try {
            Remove-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $script:TargetServicePrincipal.Id `
                -AppRoleAssignmentId $assignment.AssignmentId `
                -ErrorAction Stop

            Write-UiLog -Message "Removed successfully: $($assignment.Permission)" -Level 'SUCCESS'
            $removed++
        }
        catch {
            Write-UiLog -Message (
                "Failed to remove {0}: {1}" -f
                $assignment.Permission,
                $_.Exception.Message
            ) -Level 'ERROR'
            $failed++
        }
    }

    Refresh-CurrentAssignments

    [System.Windows.Forms.MessageBox]::Show(
        "Completed.`n`nRemoved: $removed`nFailed: $failed",
        'Graph App Role Manager',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

# ---------------------------------------------------------------------------
# FORM
# ---------------------------------------------------------------------------

$colorCanvas = [System.Drawing.Color]::FromArgb(243, 246, 250)
$colorHeader = [System.Drawing.Color]::FromArgb(28, 38, 53)
$colorBlue = [System.Drawing.Color]::FromArgb(47, 128, 237)
$colorGreen = [System.Drawing.Color]::FromArgb(22, 163, 74)
$colorRed = [System.Drawing.Color]::FromArgb(190, 45, 45)
$colorMuted = [System.Drawing.Color]::FromArgb(96, 108, 126)

function Set-PrimaryButtonStyle {
    param(
        [Parameter(Mandatory)]$Button,
        [System.Drawing.Color]$BackColor = $colorBlue
    )

    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 0
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Margin = New-Object System.Windows.Forms.Padding(6)
}

function New-FieldLabel {
    param([Parameter(Mandatory)][string]$Text)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
    $label.Anchor = 'Left'
    return $label
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Graph App Role Manager — Native Windows'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1280, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1080, 740)
$form.BackColor = $colorCanvas
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.AutoScaleMode = 'Dpi'

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = 'Fill'
$rootLayout.ColumnCount = 1
$rootLayout.RowCount = 2
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 96)))
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
$form.Controls.Add($rootLayout)

$headerPanel = New-Object System.Windows.Forms.TableLayoutPanel
$headerPanel.Dock = 'Fill'
$headerPanel.BackColor = $colorHeader
$headerPanel.ColumnCount = 2
$headerPanel.RowCount = 1
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(24, 14, 24, 14)
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
$headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 250)))
$rootLayout.Controls.Add($headerPanel, 0, 0)

$titlePanel = New-Object System.Windows.Forms.TableLayoutPanel
$titlePanel.Dock = 'Fill'
$titlePanel.ColumnCount = 1
$titlePanel.RowCount = 2
$titlePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 62)))
$titlePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 38)))
$headerPanel.Controls.Add($titlePanel, 0, 0)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Graph App Role Manager'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20)
$lblTitle.AutoSize = $true
$lblTitle.Anchor = 'Left'
$titlePanel.Controls.Add($lblTitle, 0, 0)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'Native Windows alternative · Assign Microsoft Graph application permissions'
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(195, 205, 219)
$lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$lblSubtitle.AutoSize = $true
$lblSubtitle.Anchor = 'Left'
$titlePanel.Controls.Add($lblSubtitle, 0, 1)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Connect to Microsoft Graph'
$btnConnect.Dock = 'Fill'
$btnConnect.Margin = New-Object System.Windows.Forms.Padding(6, 16, 0, 16)
Set-PrimaryButtonStyle -Button $btnConnect
$headerPanel.Controls.Add($btnConnect, 1, 0)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabs.Padding = New-Object System.Drawing.Point(18, 7)
$tabs.Margin = New-Object System.Windows.Forms.Padding(12)
$rootLayout.Controls.Add($tabs, 0, 1)

# ---------------------------------------------------------------------------
# TAB 1 - ASSIGN
# ---------------------------------------------------------------------------

$tabAssign = New-Object System.Windows.Forms.TabPage
$tabAssign.Text = 'Manage permissions'
$tabAssign.BackColor = $colorCanvas
$tabAssign.Padding = New-Object System.Windows.Forms.Padding(14)
$tabs.TabPages.Add($tabAssign)

$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 2
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 235)))
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
$tabAssign.Controls.Add($mainLayout)

$grpIdentity = New-Object System.Windows.Forms.GroupBox
$grpIdentity.Text = '1. Choose the target identity'
$grpIdentity.Dock = 'Fill'
$grpIdentity.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 14)
$grpIdentity.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$mainLayout.Controls.Add($grpIdentity, 0, 0)

$identityLayout = New-Object System.Windows.Forms.TableLayoutPanel
$identityLayout.Dock = 'Fill'
$identityLayout.ColumnCount = 3
$identityLayout.RowCount = 4
$identityLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
$identityLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 130)))
$identityLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 0)))
$identityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 27)))
$identityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 38)))
$identityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 57)))
$identityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
$grpIdentity.Controls.Add($identityLayout)

$lblIdentitySearch = New-FieldLabel -Text 'Managed identity or service-principal display name'
$identityLayout.Controls.Add($lblIdentitySearch, 0, 0)

$txtIdentityName = New-Object System.Windows.Forms.TextBox
$txtIdentityName.PlaceholderText = 'Enter a full or partial display name'
$txtIdentityName.Dock = 'Fill'
$txtIdentityName.Margin = New-Object System.Windows.Forms.Padding(0, 3, 8, 6)
$identityLayout.Controls.Add($txtIdentityName, 0, 1)

$btnSearchIdentity = New-Object System.Windows.Forms.Button
$btnSearchIdentity.Text = 'Search'
$btnSearchIdentity.Dock = 'Fill'
$btnSearchIdentity.Margin = New-Object System.Windows.Forms.Padding(4, 2, 0, 5)
Set-PrimaryButtonStyle -Button $btnSearchIdentity
$identityLayout.Controls.Add($btnSearchIdentity, 1, 1)

$resultsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$resultsLayout.Dock = 'Fill'
$resultsLayout.ColumnCount = 1
$resultsLayout.RowCount = 2
$resultsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 23)))
$resultsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
$identityLayout.SetColumnSpan($resultsLayout, 2)
$identityLayout.Controls.Add($resultsLayout, 0, 2)

$lblSearchResults = New-FieldLabel -Text 'Matching identities'
$resultsLayout.Controls.Add($lblSearchResults, 0, 0)

$cmbIdentityResults = New-Object System.Windows.Forms.ComboBox
$cmbIdentityResults.DropDownStyle = 'DropDownList'
$cmbIdentityResults.Dock = 'Fill'
$cmbIdentityResults.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
$resultsLayout.Controls.Add($cmbIdentityResults, 0, 1)

$summaryPanel = New-Object System.Windows.Forms.TableLayoutPanel
$summaryPanel.Dock = 'Fill'
$summaryPanel.BackColor = [System.Drawing.Color]::White
$summaryPanel.CellBorderStyle = 'Single'
$summaryPanel.ColumnCount = 4
$summaryPanel.RowCount = 2
$summaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 85)))
$summaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
$summaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 85)))
$summaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
$summaryPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))
$summaryPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))
$identityLayout.SetColumnSpan($summaryPanel, 2)
$identityLayout.Controls.Add($summaryPanel, 0, 3)

$lblIdentityName = New-FieldLabel -Text 'Name'
$lblIdentityNameValue = New-FieldLabel -Text '-'
$lblIdentityNameValue.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lblObjectId = New-FieldLabel -Text 'Object ID'
$lblObjectIdValue = New-FieldLabel -Text '-'
$lblObjectIdValue.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lblAppId = New-FieldLabel -Text 'App ID'
$lblAppIdValue = New-FieldLabel -Text '-'
$lblAppIdValue.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$summaryPanel.Controls.Add($lblIdentityName, 0, 0)
$summaryPanel.Controls.Add($lblIdentityNameValue, 1, 0)
$summaryPanel.Controls.Add($lblObjectId, 0, 1)
$summaryPanel.Controls.Add($lblObjectIdValue, 1, 1)
$summaryPanel.Controls.Add($lblAppId, 2, 1)
$summaryPanel.Controls.Add($lblAppIdValue, 3, 1)
$summaryPanel.SetColumnSpan($lblIdentityNameValue, 3)

$splitPermissions = New-Object System.Windows.Forms.SplitContainer
$splitPermissions.Dock = 'Fill'
$splitPermissions.SplitterDistance = 690
$splitPermissions.Panel1MinSize = 480
$splitPermissions.Panel2MinSize = 360
$mainLayout.Controls.Add($splitPermissions, 0, 1)

$grpPermissions = New-Object System.Windows.Forms.GroupBox
$grpPermissions.Text = '2. Select Microsoft Graph application permissions'
$grpPermissions.Dock = 'Fill'
$grpPermissions.Padding = New-Object System.Windows.Forms.Padding(14)
$splitPermissions.Panel1.Controls.Add($grpPermissions)

$permissionLayout = New-Object System.Windows.Forms.TableLayoutPanel
$permissionLayout.Dock = 'Fill'
$permissionLayout.ColumnCount = 2
$permissionLayout.RowCount = 3
$permissionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
$permissionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 170)))
$permissionLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 42)))
$permissionLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
$permissionLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 44)))
$grpPermissions.Controls.Add($permissionLayout)

$txtPermissionFilter = New-Object System.Windows.Forms.TextBox
$txtPermissionFilter.PlaceholderText = 'Filter permissions, e.g. DeviceManagement or Sites'
$txtPermissionFilter.Dock = 'Fill'
$txtPermissionFilter.Margin = New-Object System.Windows.Forms.Padding(0, 5, 8, 7)
$permissionLayout.Controls.Add($txtPermissionFilter, 0, 0)

$btnLoadPermissions = New-Object System.Windows.Forms.Button
$btnLoadPermissions.Text = 'Reload permissions'
$btnLoadPermissions.Dock = 'Fill'
$btnLoadPermissions.Margin = New-Object System.Windows.Forms.Padding(4, 3, 0, 6)
Set-PrimaryButtonStyle -Button $btnLoadPermissions
$permissionLayout.Controls.Add($btnLoadPermissions, 1, 0)

$checkedPermissions = New-Object System.Windows.Forms.CheckedListBox
$checkedPermissions.CheckOnClick = $true
$checkedPermissions.Dock = 'Fill'
$checkedPermissions.HorizontalScrollbar = $true
$checkedPermissions.DisplayMember = 'Text'
$checkedPermissions.BackColor = [System.Drawing.Color]::White
$permissionLayout.SetColumnSpan($checkedPermissions, 2)
$permissionLayout.Controls.Add($checkedPermissions, 0, 1)

$lblPermissionCount = New-Object System.Windows.Forms.Label
$lblPermissionCount.Text = '0 permission(s) displayed'
$lblPermissionCount.AutoSize = $true
$lblPermissionCount.ForeColor = $colorMuted
$lblPermissionCount.Anchor = 'Left'
$permissionLayout.Controls.Add($lblPermissionCount, 0, 2)

$btnAssign = New-Object System.Windows.Forms.Button
$btnAssign.Text = 'Assign selected'
$btnAssign.Dock = 'Fill'
$btnAssign.Margin = New-Object System.Windows.Forms.Padding(4, 6, 0, 0)
Set-PrimaryButtonStyle -Button $btnAssign -BackColor $colorGreen
$permissionLayout.Controls.Add($btnAssign, 1, 2)

$grpCurrent = New-Object System.Windows.Forms.GroupBox
$grpCurrent.Text = '3. Current Microsoft Graph assignments'
$grpCurrent.Dock = 'Fill'
$grpCurrent.Padding = New-Object System.Windows.Forms.Padding(14)
$splitPermissions.Panel2.Controls.Add($grpCurrent)

$assignmentLayout = New-Object System.Windows.Forms.TableLayoutPanel
$assignmentLayout.Dock = 'Fill'
$assignmentLayout.ColumnCount = 2
$assignmentLayout.RowCount = 3
$assignmentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
$assignmentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 120)))
$assignmentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 42)))
$assignmentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
$assignmentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 44)))
$grpCurrent.Controls.Add($assignmentLayout)

$lblAssignedCount = New-Object System.Windows.Forms.Label
$lblAssignedCount.Text = '0 permission(s) assigned'
$lblAssignedCount.ForeColor = $colorMuted
$lblAssignedCount.AutoSize = $true
$lblAssignedCount.Anchor = 'Left'
$assignmentLayout.Controls.Add($lblAssignedCount, 0, 0)

$btnRefreshAssignments = New-Object System.Windows.Forms.Button
$btnRefreshAssignments.Text = 'Refresh'
$btnRefreshAssignments.Dock = 'Fill'
$btnRefreshAssignments.Margin = New-Object System.Windows.Forms.Padding(4, 3, 0, 6)
Set-PrimaryButtonStyle -Button $btnRefreshAssignments
$assignmentLayout.Controls.Add($btnRefreshAssignments, 1, 0)

$gridAssignments = New-Object System.Windows.Forms.DataGridView
$gridAssignments.Dock = 'Fill'
$gridAssignments.AllowUserToAddRows = $false
$gridAssignments.AllowUserToDeleteRows = $false
$gridAssignments.AllowUserToResizeRows = $false
$gridAssignments.ReadOnly = $true
$gridAssignments.RowHeadersVisible = $false
$gridAssignments.SelectionMode = 'FullRowSelect'
$gridAssignments.MultiSelect = $true
$gridAssignments.AutoSizeColumnsMode = 'Fill'
$gridAssignments.BackgroundColor = [System.Drawing.Color]::White
$gridAssignments.BorderStyle = 'Fixed3D'
$gridAssignments.EnableHeadersVisualStyles = $false
$gridAssignments.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(229, 235, 244)
$gridAssignments.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
$gridAssignments.DefaultCellStyle.SelectionBackColor = $colorBlue
$assignmentLayout.SetColumnSpan($gridAssignments, 2)
$assignmentLayout.Controls.Add($gridAssignments, 0, 1)
[void]$gridAssignments.Columns.Add('Permission', 'Permission')
[void]$gridAssignments.Columns.Add('DisplayName', 'Display name')
[void]$gridAssignments.Columns.Add('RoleId', 'App role ID')
[void]$gridAssignments.Columns.Add('AssignmentId', 'Assignment ID')
$gridAssignments.Columns['RoleId'].Visible = $false
$gridAssignments.Columns['AssignmentId'].Visible = $false

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = 'Remove selected'
$btnRemove.Dock = 'Fill'
$btnRemove.Margin = New-Object System.Windows.Forms.Padding(4, 6, 0, 0)
Set-PrimaryButtonStyle -Button $btnRemove -BackColor $colorRed
$assignmentLayout.Controls.Add($btnRemove, 1, 2)

$tabLog = New-Object System.Windows.Forms.TabPage
$tabLog.Text = 'Activity log'
$tabLog.Padding = New-Object System.Windows.Forms.Padding(12)
$tabs.TabPages.Add($tabLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Dock = 'Fill'
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Both'
$txtLog.WordWrap = $false
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(17, 24, 39)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(229, 231, 235)
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabLog.Controls.Add($txtLog)

# ---------------------------------------------------------------------------
# EVENTS
# ---------------------------------------------------------------------------

$btnConnect.Add_Click({
    try {
        Set-BusyState -Busy $true
        Ensure-GraphModules

        Write-UiLog -Message 'Opening Microsoft Graph authentication...'

        Connect-MgGraph `
            -Scopes @(
                'Application.Read.All',
                'AppRoleAssignment.ReadWrite.All'
            ) `
            -NoWelcome `
            -ContextScope Process `
            -ErrorAction Stop

        $context = Get-MgContext
        $script:GraphConnected = $true

        $btnConnect.Text = "Connected: $($context.Account)"
        $btnConnect.BackColor = [System.Drawing.Color]::FromArgb(22, 163, 74)

        Write-UiLog -Message (
            "Connected to tenant {0} as {1}." -f
            $context.TenantId,
            $context.Account
        ) -Level 'SUCCESS'

        Load-GraphApplicationRoles
    }
    catch {
        Show-UiError -Message $_.Exception.Message
    }
    finally {
        Set-BusyState -Busy $false
    }
})

$btnSearchIdentity.Add_Click({
    try {
        Set-BusyState -Busy $true
        Search-TargetServicePrincipals
    }
    catch {
        Show-UiError -Message $_.Exception.Message
    }
    finally {
        Set-BusyState -Busy $false
    }
})

$cmbIdentityResults.Add_SelectedIndexChanged({
    try {
        $selected = Get-SelectedTargetServicePrincipal
        $script:TargetServicePrincipal = $selected
        Update-IdentitySummary -ServicePrincipal $selected

        if ($null -ne $selected) {
            Write-UiLog -Message "Selected target: $($selected.DisplayName)"
            Refresh-CurrentAssignments
        }
    }
    catch {
        Show-UiError -Message $_.Exception.Message
    }
})

$btnLoadPermissions.Add_Click({
    try {
        Set-BusyState -Busy $true
        Load-GraphApplicationRoles
    }
    catch {
        Show-UiError -Message $_.Exception.Message
    }
    finally {
        Set-BusyState -Busy $false
    }
})

$txtPermissionFilter.Add_TextChanged({
    try {
        Apply-PermissionFilter
    }
    catch {
        Write-UiLog -Message $_.Exception.Message -Level 'ERROR'
    }
})

$btnRefreshAssignments.Add_Click({
    try {
        Set-BusyState -Busy $true
        Refresh-CurrentAssignments
    }
    catch {
        Show-UiError -Message $_.Exception.Message
    }
    finally {
        Set-BusyState -Busy $false
    }
})

$btnAssign.Add_Click({
    try {
        Set-BusyState -Busy $true
        Assign-SelectedPermissions
    }
    catch {
        Show-UiError -Message $_.Exception.Message
    }
    finally {
        Set-BusyState -Busy $false
    }
})

$btnRemove.Add_Click({
    try {
        Set-BusyState -Busy $true
        Remove-SelectedPermissions
    }
    catch {
        Show-UiError -Message $_.Exception.Message
    }
    finally {
        Set-BusyState -Busy $false
    }
})

$form.Add_FormClosing({
    try {
        if ($script:GraphConnected) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
    catch {}
})

$form.Add_Shown({
    Write-UiLog -Message 'Graph App Role Manager started.'
    Write-UiLog -Message 'Connect to Microsoft Graph, search the Managed Identity, then select application permissions.'
})

# ---------------------------------------------------------------------------
# START
# ---------------------------------------------------------------------------

if (-not $NoGui) {
    [void]$form.ShowDialog()
}
