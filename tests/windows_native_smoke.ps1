$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Assert-True {
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$guiPath = Join-Path $repositoryRoot 'windows/Graph-App-Role-Manager.ps1'
$guiSource = Get-Content -LiteralPath $guiPath -Raw

# Load and construct the complete form without showing it.
. $guiPath -NoGui

try {
    Assert-True ($form.Text -like '*Native Windows*') 'The native Windows form was not constructed.'
    Assert-True ($rootLayout -is [System.Windows.Forms.TableLayoutPanel]) 'The responsive root layout is missing.'
    Assert-True ($btnConnect.Text -eq 'Connect to Microsoft Graph') 'The Graph connection action is missing.'
    Assert-True ($btnConnect.Height -eq 40) 'The compact Graph connection button is missing.'
    Assert-True ($btnConnect.FlatStyle -eq 'Flat') 'The modern button styling is missing.'
    Assert-True ($null -ne $btnConnect.Region) 'The rounded button styling is missing.'
    Assert-True ($tabs.TabPages.Count -eq 2) 'The permissions and activity-log tabs were not created.'
    Assert-True ($btnSearchIdentity.Text -eq 'Clear') 'The identity filter clear action is missing.'

    # Regression check for the StrictMode failure reported when exactly one
    # Microsoft Graph module is missing.
    Assert-True (
        $guiSource -match '(?s)\$missingModules\s*=\s*@\(\s*foreach'
    ) 'The missing-module result must remain array-wrapped.'
    Assert-True (
        $guiSource -match '(?s)Load-GraphApplicationRoles\s+Load-TargetServicePrincipals'
    ) 'Identities must load automatically after connection.'
    Assert-True (
        $guiSource -match 'ClientSize\.Width\s*-\s*\$splitPermissions\.SplitterWidth\)\s*/\s*2'
    ) 'The permission and assignment panels must remain balanced.'
}
finally {
    $form.Dispose()
}

Write-Host 'Native Windows GUI smoke test passed.'
