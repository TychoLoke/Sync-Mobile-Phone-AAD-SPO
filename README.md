# Sync Mobile Phone from Azure Active Directory to SharePoint Online using PowerShell
This script is used to synchronize the mobile phone number from Azure AD to SharePoint Online user profile property. It imports the MSOnline and Microsoft.Online.SharePoint.PowerShell modules, adds SharePoint CSOM libraries, sets defaults, and prompts for Azure AD and SharePoint Online admin credentials. It then connects to Azure AD, gets all Azure AD users, and for each user, it checks to see if the Azure AD user has a mobile phone number specified. If so, it gets the existing value of the SharePoint Online user profile property for "CellPhone" and updates it with the mobile phone number from Azure AD.

## Requirements
- Sharepoint Online
- Azure AD
- SharePoint Online Management Shell
- Admin Account without MFA and the Sharepoint and Service Admin role

## Setup
1. Make sure the SharePoint Online Management Shell is installed on the machine where the script will be executed.
2. Update the $spoAdminUrl variable with the SharePoint Online admin center URL of your tenant.
3. Update the path of the SharePoint CSOM libraries to match the location on your machine.
4. Run the script.

## Usage
1. Run the script with PowerShell with Admin Rights
2. Enter the credentials for the Azure AD and SharePoint Online admin when prompted.
3. The script will then synchronize the mobile phone number from Azure AD to SharePoint Online user profile property for all users.

## Note
- By default, the script does not overwrite existing SharePoint Online user profile property values. If you want to overwrite existing values, set the $overwriteExistingSPOUPAValue variable to $true.
- If Azure AD user's mobile phone number is not specified, the script will skip that user and print a message to the console.

## Troubleshooting

If you encounter any issues, check the error message and the PowerShell console for more details.

Please report any issues or suggestions on the repository's issue tracker.
