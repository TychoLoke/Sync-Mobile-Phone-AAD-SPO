# Sync Mobile Phone from Microsoft Entra ID to SharePoint Online using PowerShell

This script synchronizes the mobile phone number from Microsoft Entra ID to the SharePoint Online user profile property. It uses Microsoft Graph for user data and `PnP.PowerShell` for SharePoint profile updates.

## Requirements

- SharePoint Online
- Microsoft Entra ID
- Permission to install PowerShell modules for the current user
- Admin account with the SharePoint and service admin roles
- Internet access the first time you run the script so it can bootstrap `PowerShellAdminHelpers`

## Setup

1. Install the required modules:
   ```powershell
   Invoke-WebRequest `
     -Uri "https://raw.githubusercontent.com/TychoLoke/powershell-admin-helpers/main/Install-PowerShellAdminHelpers.ps1" `
     -OutFile "$env:TEMP\Install-PowerShellAdminHelpers.ps1"

   & "$env:TEMP\Install-PowerShellAdminHelpers.ps1"
   ```
2. Run the script with your tenant-specific values:

   ```powershell
   .\sync-mobile-phone-aad-spo.ps1 `
     -TenantId "<tenant-id>" `
     -SpoAdminUrl "https://contoso-admin.sharepoint.com" `
     -OverwriteExistingSPOUPAValue $true
   ```

**Usage:**

1. Run the script from an elevated PowerShell session.
2. You will be prompted to authenticate to Microsoft Graph interactively.
3. You will be prompted to log in to SharePoint Online.
4. The script will then synchronize the mobile phone number from Entra ID to the SharePoint Online user profile property for all users.

**Notes:**

- If the Entra ID user's mobile phone number is not specified, the script will skip that user and print a message to the console.
- By default, the script does not overwrite existing SharePoint Online user profile property values. If you want to overwrite existing values, set the `$overwriteExistingSPOUPAValue` variable to `$true`.

**Troubleshooting:**

If you encounter any issues, check the error message and the PowerShell console for more details.

Please report any issues or suggestions on the repository's issue tracker.

## Notes

- Test the script in a non-production tenant before using it against a live directory.
- The script uses Microsoft Graph delegated permissions and interactive PnP authentication.
- The script installs `Microsoft.Graph.Users` and `PnP.PowerShell` automatically if they are missing.

**Changelog:**

- **Version 2.1.0 (Latest)**
  - Updated to use Microsoft Graph and `PnP.PowerShell`.
  - Replaced legacy AzureAD and `SharePointPnPPowerShellOnline` module usage.

- **Version 1.0.0**
  - Initial release.
