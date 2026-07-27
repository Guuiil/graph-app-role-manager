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

    $missingModules = foreach ($moduleName in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            $moduleName
        }
    }

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

        $comboItem = New-Object System.Windows.Forms.ComboBox
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

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Graph App Role Manager'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1180, 820)
$form.MinimumSize = New-Object System.Drawing.Size(1050, 720)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.AutoScaleMode = 'Dpi'

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Top'
$headerPanel.Height = 86
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
$form.Controls.Add($headerPanel)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Graph App Role Manager'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(24, 14)
$headerPanel.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'Assign Microsoft Graph application permissions to a Managed Identity'
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(209, 213, 219)
$lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$lblSubtitle.AutoSize = $true
$lblSubtitle.Location = New-Object System.Drawing.Point(27, 52)
$headerPanel.Controls.Add($lblSubtitle)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Connect to Microsoft Graph'
$btnConnect.Size = New-Object System.Drawing.Size(220, 36)
$btnConnect.Anchor = 'Top,Right'
$btnConnect.Location = New-Object System.Drawing.Point(930, 24)
$btnConnect.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
$btnConnect.ForeColor = [System.Drawing.Color]::White
$btnConnect.FlatStyle = 'Flat'
$btnConnect.FlatAppearance.BorderSize = 0
$headerPanel.Controls.Add($btnConnect)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabs.Padding = New-Object System.Drawing.Point(18, 7)
$form.Controls.Add($tabs)
$tabs.BringToFront()

# ---------------------------------------------------------------------------
# TAB 1 - ASSIGN
# ---------------------------------------------------------------------------

$tabAssign = New-Object System.Windows.Forms.TabPage
$tabAssign.Text = 'Assign permissions'
$tabAssign.BackColor = $form.BackColor
$tabAssign.Padding = New-Object System.Windows.Forms.Padding(18)
$tabs.TabPages.Add($tabAssign)

$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock = 'Fill'
$splitMain.Orientation = 'Horizontal'
$splitMain.SplitterDistance = 250
$splitMain.IsSplitterFixed = $false
$tabAssign.Controls.Add($splitMain)

# Identity group
$grpIdentity = New-Object System.Windows.Forms.GroupBox
$grpIdentity.Text = '1. Select the target Managed Identity'
$grpIdentity.Dock = 'Fill'
$grpIdentity.Padding = New-Object System.Windows.Forms.Padding(14)
$splitMain.Panel1.Controls.Add($grpIdentity)

$lblIdentitySearch = New-Object System.Windows.Forms.Label
$lblIdentitySearch.Text = 'Display name'
$lblIdentitySearch.AutoSize = $true
$lblIdentitySearch.Location = New-Object System.Drawing.Point(18, 35)
$grpIdentity.Controls.Add($lblIdentitySearch)

$txtIdentityName = New-Object System.Windows.Forms.TextBox
$txtIdentityName.Text = 'AA-ModernWorkplace-Reporting'
$txtIdentityName.Location = New-Object System.Drawing.Point(18, 58)
$txtIdentityName.Size = New-Object System.Drawing.Size(520, 27)
$grpIdentity.Controls.Add($txtIdentityName)

$btnSearchIdentity = New-Object System.Windows.Forms.Button
$btnSearchIdentity.Text = 'Search'
$btnSearchIdentity.Location = New-Object System.Drawing.Point(550, 56)
$btnSearchIdentity.Size = New-Object System.Drawing.Size(105, 30)
$grpIdentity.Controls.Add($btnSearchIdentity)

$lblSearchResults = New-Object System.Windows.Forms.Label
$lblSearchResults.Text = 'Matching service principals'
$lblSearchResults.AutoSize = $true
$lblSearchResults.Location = New-Object System.Drawing.Point(18, 98)
$grpIdentity.Controls.Add($lblSearchResults)

$cmbIdentityResults = New-Object System.Windows.Forms.ComboBox
$cmbIdentityResults.DropDownStyle = 'DropDownList'
$cmbIdentityResults.Location = New-Object System.Drawing.Point(18, 121)
$cmbIdentityResults.Size = New-Object System.Drawing.Size(850, 28)
$cmbIdentityResults.Anchor = 'Top,Left,Right'
$grpIdentity.Controls.Add($cmbIdentityResults)

$summaryPanel = New-Object System.Windows.Forms.Panel
$summaryPanel.Location = New-Object System.Drawing.Point(18, 162)
$summaryPanel.Size = New-Object System.Drawing.Size(1080, 67)
$summaryPanel.Anchor = 'Top,Left,Right'
$summaryPanel.BackColor = [System.Drawing.Color]::White
$summaryPanel.BorderStyle = 'FixedSingle'
$grpIdentity.Controls.Add($summaryPanel)

$lblIdentityName = New-Object System.Windows.Forms.Label
$lblIdentityName.Text = 'Name:'
$lblIdentityName.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$lblIdentityName.Location = New-Object System.Drawing.Point(12, 9)
$lblIdentityName.AutoSize = $true
$summaryPanel.Controls.Add($lblIdentityName)

$lblIdentityNameValue = New-Object System.Windows.Forms.Label
$lblIdentityNameValue.Text = '-'
$lblIdentityNameValue.Location = New-Object System.Drawing.Point(80, 9)
$lblIdentityNameValue.AutoSize = $true
$summaryPanel.Controls.Add($lblIdentityNameValue)

$lblObjectId = New-Object System.Windows.Forms.Label
$lblObjectId.Text = 'Object ID:'
$lblObjectId.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$lblObjectId.Location = New-Object System.Drawing.Point(12, 35)
$lblObjectId.AutoSize = $true
$summaryPanel.Controls.Add($lblObjectId)

$lblObjectIdValue = New-Object System.Windows.Forms.Label
$lblObjectIdValue.Text = '-'
$lblObjectIdValue.Location = New-Object System.Drawing.Point(80, 35)
$lblObjectIdValue.AutoSize = $true
$summaryPanel.Controls.Add($lblObjectIdValue)

$lblAppId = New-Object System.Windows.Forms.Label
$lblAppId.Text = 'App ID:'
$lblAppId.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$lblAppId.Location = New-Object System.Drawing.Point(570, 35)
$lblAppId.AutoSize = $true
$summaryPanel.Controls.Add($lblAppId)

$lblAppIdValue = New-Object System.Windows.Forms.Label
$lblAppIdValue.Text = '-'
$lblAppIdValue.Location = New-Object System.Drawing.Point(630, 35)
$lblAppIdValue.AutoSize = $true
$summaryPanel.Controls.Add($lblAppIdValue)

# Bottom permissions area
$splitPermissions = New-Object System.Windows.Forms.SplitContainer
$splitPermissions.Dock = 'Fill'
$splitPermissions.SplitterDistance = 610
$splitPermissions.IsSplitterFixed = $false
$splitMain.Panel2.Controls.Add($splitPermissions)

$grpPermissions = New-Object System.Windows.Forms.GroupBox
$grpPermissions.Text = '2. Select Microsoft Graph application permissions'
$grpPermissions.Dock = 'Fill'
$grpPermissions.Padding = New-Object System.Windows.Forms.Padding(14)
$splitPermissions.Panel1.Controls.Add($grpPermissions)

$txtPermissionFilter = New-Object System.Windows.Forms.TextBox
$txtPermissionFilter.PlaceholderText = 'Filter permissions, e.g. DeviceManagement or Sites'
$txtPermissionFilter.Location = New-Object System.Drawing.Point(18, 32)
$txtPermissionFilter.Size = New-Object System.Drawing.Size(410, 27)
$txtPermissionFilter.Anchor = 'Top,Left,Right'
$grpPermissions.Controls.Add($txtPermissionFilter)

$btnLoadPermissions = New-Object System.Windows.Forms.Button
$btnLoadPermissions.Text = 'Load permissions'
$btnLoadPermissions.Location = New-Object System.Drawing.Point(440, 30)
$btnLoadPermissions.Size = New-Object System.Drawing.Size(140, 30)
$btnLoadPermissions.Anchor = 'Top,Right'
$grpPermissions.Controls.Add($btnLoadPermissions)

$checkedPermissions = New-Object System.Windows.Forms.CheckedListBox
$checkedPermissions.CheckOnClick = $true
$checkedPermissions.Location = New-Object System.Drawing.Point(18, 72)
$checkedPermissions.Size = New-Object System.Drawing.Size(562, 310)
$checkedPermissions.Anchor = 'Top,Bottom,Left,Right'
$checkedPermissions.HorizontalScrollbar = $true
$checkedPermissions.DisplayMember = 'Text'
$grpPermissions.Controls.Add($checkedPermissions)

$lblPermissionCount = New-Object System.Windows.Forms.Label
$lblPermissionCount.Text = '0 permission(s) displayed'
$lblPermissionCount.AutoSize = $true
$lblPermissionCount.Location = New-Object System.Drawing.Point(18, 390)
$lblPermissionCount.Anchor = 'Bottom,Left'
$grpPermissions.Controls.Add($lblPermissionCount)

$btnAssign = New-Object System.Windows.Forms.Button
$btnAssign.Text = 'Assign selected permissions'
$btnAssign.Location = New-Object System.Drawing.Point(365, 384)
$btnAssign.Size = New-Object System.Drawing.Size(215, 34)
$btnAssign.Anchor = 'Bottom,Right'
$btnAssign.BackColor = [System.Drawing.Color]::FromArgb(22, 163, 74)
$btnAssign.ForeColor = [System.Drawing.Color]::White
$btnAssign.FlatStyle = 'Flat'
$btnAssign.FlatAppearance.BorderSize = 0
$grpPermissions.Controls.Add($btnAssign)

$grpCurrent = New-Object System.Windows.Forms.GroupBox
$grpCurrent.Text = '3. Current Microsoft Graph assignments'
$grpCurrent.Dock = 'Fill'
$grpCurrent.Padding = New-Object System.Windows.Forms.Padding(14)
$splitPermissions.Panel2.Controls.Add($grpCurrent)

$btnRefreshAssignments = New-Object System.Windows.Forms.Button
$btnRefreshAssignments.Text = 'Refresh'
$btnRefreshAssignments.Location = New-Object System.Drawing.Point(18, 30)
$btnRefreshAssignments.Size = New-Object System.Drawing.Size(100, 30)
$grpCurrent.Controls.Add($btnRefreshAssignments)

$lblAssignedCount = New-Object System.Windows.Forms.Label
$lblAssignedCount.Text = '0 Microsoft Graph permission(s) assigned'
$lblAssignedCount.AutoSize = $true
$lblAssignedCount.Location = New-Object System.Drawing.Point(132, 37)
$grpCurrent.Controls.Add($lblAssignedCount)

$gridAssignments = New-Object System.Windows.Forms.DataGridView
$gridAssignments.Location = New-Object System.Drawing.Point(18, 72)
$gridAssignments.Size = New-Object System.Drawing.Size(465, 304)
$gridAssignments.Anchor = 'Top,Bottom,Left,Right'
$gridAssignments.AllowUserToAddRows = $false
$gridAssignments.AllowUserToDeleteRows = $false
$gridAssignments.AllowUserToResizeRows = $false
$gridAssignments.ReadOnly = $true
$gridAssignments.RowHeadersVisible = $false
$gridAssignments.SelectionMode = 'FullRowSelect'
$gridAssignments.AutoSizeColumnsMode = 'Fill'
$gridAssignments.BackgroundColor = [System.Drawing.Color]::White
[void]$gridAssignments.Columns.Add('Permission', 'Permission')
[void]$gridAssignments.Columns.Add('DisplayName', 'Display name')
[void]$gridAssignments.Columns.Add('RoleId', 'App role ID')
[void]$gridAssignments.Columns.Add('AssignmentId', 'Assignment ID')
$gridAssignments.Columns['RoleId'].Visible = $false
$gridAssignments.Columns['AssignmentId'].Visible = $false
$grpCurrent.Controls.Add($gridAssignments)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = 'Remove selected'
$btnRemove.Location = New-Object System.Drawing.Point(303, 384)
$btnRemove.Size = New-Object System.Drawing.Size(180, 34)
$btnRemove.Anchor = 'Bottom,Right'
$btnRemove.BackColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
$btnRemove.ForeColor = [System.Drawing.Color]::White
$btnRemove.FlatStyle = 'Flat'
$btnRemove.FlatAppearance.BorderSize = 0
$grpCurrent.Controls.Add($btnRemove)

# ---------------------------------------------------------------------------
# TAB 2 - LOG
# ---------------------------------------------------------------------------

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

[void]$form.ShowDialog()
