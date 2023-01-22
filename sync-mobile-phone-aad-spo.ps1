Import-Module MSOnline
Import-Module Microsoft.Online.SharePoint.PowerShell

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
    Connect-MsolService -Credential $credential

    # Get credentials for SharePointOnline
    $spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credential.UserName, $credential.Password)
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($spoAdminUrl)
    $ctx.Credentials = $spoCredentials
    $spoPeopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($ctx)

    # Get all AzureAD Users
    $AzureADUsers = Get-MsolUser -All

    ForEach ($AzureADUser in $AzureADUsers) {

        $mobilePhone = $AzureADUser.MobilePhone
        $targetUPN = $AzureADUser.UserPrincipalName

        # Check to see if the AzureAD User has a MobilePhone specified
        if ($mobilePhone) {
            # Get the existing value of the SPO User Profile Property CellPhone
            $targetUserCellPhone = $spoPeopleManager.GetUserProfilePropertyFor($targetUPN, "CellPhone")
            $ctx.ExecuteQuery()

            $userCellPhone = $targetUserCellPhone.Value

            # If target property is empty let's populate it
            if (!$userCellPhone) {
                $spoPeopleManager.SetSingleValueProfileProperty($targetUPN, "CellPhone", $mobilePhone)
                $ctx.ExecuteQuery()
            }
            else {
                # Target property is not empty
                # Check to see if we're to overwrite existing property value
                if ($overwriteExistingSPOUPAValue) {
                    $spoPeopleManager.SetSingleValueProfileProperty($targetUPN, "CellPhone", $mobilePhone)
                    $ctx.ExecuteQuery()
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
