# Sync Mobile Phone from Azure Active Directory to SharePoint Online using PowerShell

This script is used to synchronize the mobile phone number from Azure AD to SharePoint Online user profile property. It imports the AzureAD and SharePointPnPPowerShellOnline modules, connects to Azure AD and SharePoint Online, and updates user profile properties based on mobile phone information from Azure AD.

## Requirements

- Sharepoint Online
- Azure AD
- SharePointPnPPowerShellOnline Module
- AzureAD PowerShell Module
- Admin Account with the SharePoint and Service Admin role (MFA supported)

## Setup

1. Install the SharePointPnPPowerShellOnline and AzureAD modules:
   ```powershell
   Install-Module -Name SharePointPnPPowerShellOnline
   Install-Module -Name AzureAD
2. Update the following variables in the script:
- `$tenantId`: Replace with your Azure AD tenant ID.
- `$spoAdminUrl`: Replace with the SharePoint Online admin center URL of your tenant.
- `$overwriteExistingSPOUPAValue`: Set to `$true` if you want to overwrite existing SharePoint Online user profile property values.
3. Run the script.

**Usage:**

1. Run the script with PowerShell with Admin Rights.
2. You will be prompted to authenticate to Azure AD interactively.
3. You will be prompted to log in to SharePoint Online.
4. The script will then synchronize the mobile phone number from Azure AD to SharePoint Online user profile property for all users.

**Notes:**

- If the Azure AD user's mobile phone number is not specified, the script will skip that user and print a message to the console.
- By default, the script does not overwrite existing SharePoint Online user profile property values. If you want to overwrite existing values, set the `$overwriteExistingSPOUPAValue` variable to `$true`.

**Troubleshooting:**

If you encounter any issues, check the error message and the PowerShell console for more details.

Please report any issues or suggestions on the repository's issue tracker.

**Changelog:**

- **Version 2.0.0 (Latest)**
  - Completely revised script to use AzureAD and SharePointPnPPowerShellOnline modules for better compatibility and ease of use.

- **Version 1.0.0**
  - Initial release.
