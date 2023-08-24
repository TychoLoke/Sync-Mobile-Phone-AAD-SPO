# Sync Mobile Phone from Azure Active Directory to SharePoint Online using PowerShell

This script is used to synchronize the mobile phone number from Azure AD to SharePoint Online user profile property. It imports the AzureAD and PnP.PowerShell modules, adds SharePoint CSOM libraries, sets defaults, and prompts for Azure AD and SharePoint Online admin credentials. It then connects to Azure AD, gets all Azure AD users, and for each user, it checks to see if the Azure AD user has a mobile phone number specified. If so, it gets the existing value of the SharePoint Online user profile property for "CellPhone" and updates it with the mobile phone number from Azure AD.

## Requirements

- Sharepoint Online
- Azure AD
- PnP.PowerShell Module
- AzureAD PowerShell Module
- Admin Account with the SharePoint and Service Admin role (MFA supported)

## Setup

1. Install the PnP.PowerShell and AzureAD modules:
   ```powershell
   Install-Module -Name PnP.PowerShell
   Install-Module -Name AzureAD
2. Update the $spoAdminUrl variable with the SharePoint Online admin center URL of your tenant.
3. Update the path of the SharePoint CSOM libraries to match the location on your machine.
4. Run the script.

## Usage

- Run the script with PowerShell with Admin Rights.
- Enter the credentials for the Azure AD and SharePoint Online admin when prompted.
- The script will then synchronize the mobile phone number from Azure AD to SharePoint Online user profile property for all users.

## Note

- By default, the script does not overwrite existing SharePoint Online user profile property values. If you want to overwrite existing values, set the $overwriteExistingSPOUPAValue variable to $true.
- If the Azure AD user's mobile phone number is not specified, the script will skip that user and print a message to the console.

## Troubleshooting

If you encounter any issues, check the error message and the PowerShell console for more details.

Please report any issues or suggestions on the repository's issue tracker.

## Changelog

### Version 2.0.0 (Latest)
- Replaced MSOnline module with AzureAD module.
- Replaced SharePoint Online credentials handling with PnP.PowerShell module.
- Simplified SharePoint Online user profile property retrieval and setting using PnP.PowerShell.

### Version 1.0.0
- Initial release with MSOnline and SharePoint CSOM libraries.
