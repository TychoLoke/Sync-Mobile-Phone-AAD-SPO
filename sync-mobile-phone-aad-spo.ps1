# Import Modules
Import-Module AzureAD
Import-Module PnP.PowerShell

# Add SharePoint CSOM libraries
Add-Type -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Add-Type -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
Add-Type -Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll'

# Defaults
$spoAdminUrl = "https://tenant-admin.sharepoint.com"
$overwriteExistingSPOUPAValue = $false

# Get credentials of account that is AzureAD Admin and SharePoint Online Admin
$credential = Get-Credential

Try {
    # Connect to AzureAD
    Connect-AzureAD -Credential $credential

    # Connect to SharePointOnline
    Connect-PnPOnline -Url $spoAdminUrl -Credentials $credential

    # Get all AzureAD Users
    $AzureADUsers = Get-AzureADUser -All $true

    ForEach ($AzureADUser in $AzureADUsers) {

        $mobilePhone = $AzureADUser.MobilePhone
        $targetUPN = $AzureADUser.UserPrincipalName

        # Check to see if the AzureAD User has a MobilePhone specified
        if ($mobilePhone) {
            # Get the existing value of the SPO User Profile Property CellPhone
            $targetUserCellPhone = Get-PnPUserProfileProperty -Account $targetUPN -PropertyName "CellPhone"

            # If target property is empty let's populate it
            if (!$targetUserCellPhone) {
                Set-PnPUserProfileProperty -Account $targetUPN -PropertyName "CellPhone" -Value $mobilePhone
            }
            else {
                # Target property is not empty
                # Check to see if we're to overwrite existing property value
                if ($overwriteExistingSPOUPAValue) {
                    Set-PnPUserProfileProperty -Account $targetUPN -PropertyName "CellPhone" -Value $mobilePhone
                }
                else {
                    # Not going to overwrite existing property value
                    Write-Output "Target SPO UPA CellPhone is not empty for $targetUPN and we're to preserve existing properties"
                }
            }
        }
        else {
            # AzureAD User MobilePhone is empty, nothing to do here
            Write-Output "AzureAD MobilePhone Property is Null or Empty for $targetUPN"
        }
    }
}
Catch {
    Write-Error $_.Exception
}
