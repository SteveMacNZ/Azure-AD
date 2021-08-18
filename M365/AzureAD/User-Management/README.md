# User Management Scripts

## Invoke-CreateCloudAccts
 Creates cloud only accounts in a M365 Tenant uses CloudUser.xlsx as a template with CSVData-Usr coped to new workbook and saved as CSV file. Currently uses MsolService

 - [ ] To do: Update script to use AzureAD powershell commands

## Invoke-MsolRestoreUser
 Restores a previously synced user account and converts to a cloud only account, uses MsolService

## Invoke-RestoreUser
 Updated version of MsolRestoreUser script that uses the AzureAD cmdlets for resetting the password and removing the ImmutableID. connects to both AzureAD and MsolService as there is currenly no native AzureAD command for restoring a deleted user. Refer to https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/users-bulk-restore for information on bulk native restore

## Invoke-LicenseUser

- [ ] To do: script to be updated and uploaded

## Invoke-AddSSPR
 Adds a list of users from a CSV file to a defined Self Service Password Reset group in Azure AD
- [X] To do: script to be updated and uploaded