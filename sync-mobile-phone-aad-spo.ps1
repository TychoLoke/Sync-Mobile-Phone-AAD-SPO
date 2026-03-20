param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$SpoAdminUrl,

    [bool]$OverwriteExistingSPOUPAValue = $true
)

function Initialize-PowerShellAdminHelpers {
    $moduleName = "PowerShellAdminHelpers"

    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        $installerPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "Install-PowerShellAdminHelpers.ps1"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/TychoLoke/powershell-admin-helpers/main/Install-PowerShellAdminHelpers.ps1" -OutFile $installerPath
        & $installerPath
    }

    Import-Module -Name $moduleName -Force -ErrorAction Stop
}

try {
    Initialize-PowerShellAdminHelpers
    Ensure-Module -ModuleName "Microsoft.Graph.Users"
    Ensure-Module -ModuleName "PnP.PowerShell"
    Connect-GraphWithScopes -Scopes @("User.Read.All") -TenantId $TenantId
    Connect-PnPOnline -Url $SpoAdminUrl -Interactive

    $entraUsers = Get-MgUser -All -Property "displayName,userPrincipalName,mobilePhone"

    foreach ($entraUser in $entraUsers) {
        $mobilePhone = $entraUser.MobilePhone
        $targetUPN = $entraUser.UserPrincipalName

        if (-not $targetUPN) {
            continue
        }

        if ($mobilePhone) {
            $userProfileProperties = Get-PnPUserProfileProperty -Account $targetUPN
            $targetUserCellPhone = $userProfileProperties.UserProfileProperties["CellPhone"]

            if (-not $targetUserCellPhone -or $OverwriteExistingSPOUPAValue) {
                Set-PnPUserProfileProperty -Account $targetUPN -PropertyName "CellPhone" -Value $mobilePhone
                Write-Output "Updated CellPhone for $targetUPN"
            }
        } else {
            Write-Output "Mobile phone is null or empty for $targetUPN"
        }
    }
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    if (Get-MgContext) {
        Disconnect-MgGraph | Out-Null
    }
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
