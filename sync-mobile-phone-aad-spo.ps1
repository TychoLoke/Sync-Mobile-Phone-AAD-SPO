param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$SpoAdminUrl,

    [bool]$OverwriteExistingSPOUPAValue = $true
)

function Ensure-Module {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Install-Module -Name $ModuleName -Scope CurrentUser -Force
    }

    Import-Module $ModuleName -ErrorAction Stop
}

try {
    Ensure-Module -ModuleName "Microsoft.Graph.Users"
    Ensure-Module -ModuleName "PnP.PowerShell"

    Connect-MgGraph -TenantId $TenantId -Scopes @("User.Read.All") -NoWelcome
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
    Disconnect-MgGraph | Out-Null
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
