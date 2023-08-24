# Powershell Script Sync Mobile Phone AAD SPO
# Author: Tycho Loke

# Import Modules
Import-Module AzureAD
Import-Module SharePointPnPPowerShellOnline

# Azure AD and SharePoint Online Connection Information
$tenantId = "2331f3d7-9649-45d8-b4c3-63bc7c48ad7f"
$spoAdminUrl = "https://peoplerockprep-admin.sharepoint.com/"
$overwriteExistingSPOUPAValue = $true

Try {
    # Authenticate to Azure AD interactively
    Connect-AzureAD -TenantId $tenantId
    
    # Connect to SharePoint Online (it will prompt for credentials)
    Connect-PnPOnline -Url $spoAdminUrl -UseWebLogin
    
    # Get all AzureAD Users
    $AzureADUsers = Get-AzureADUser -All $true
    
    ForEach ($AzureADUser in $AzureADUsers) {
        $mobilePhone = $AzureADUser.Mobile
        $targetUPN = $AzureADUser.UserPrincipalName
        
        if ($mobilePhone) {
            # Get the existing user profile properties
            $userProfileProperties = Get-PnPUserProfileProperty -Account $targetUPN
            # Get the current value of the CellPhone property
            $targetUserCellPhone = $userProfileProperties.UserProfileProperties["CellPhone"]
            
            # If target property is empty or overwrite is allowed, then update it
            if (!$targetUserCellPhone -or $overwriteExistingSPOUPAValue) {
                Set-PnPUserProfileProperty -Account $targetUPN -PropertyName "CellPhone" -Value $mobilePhone
            }
        }
        else {
            Write-Output "AzureAD MobilePhone Property is Null or Empty for $targetUPN"
        }
    }
}
Catch {
    Write-Error $_.Exception
}

# End of script
