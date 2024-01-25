<#
.SYNOPSIS
  Creates Dynamic Azure AD Mail enabled security groups and assigned to PIM roles
.DESCRIPTION
  Creates Dynamic Azure AD Mail enabled security groups and assigned to PIM roles. This feauture was still in preview confirm before production use
  ! To work on configuring of the PIM role eilgability and activation settings still to be done
  https://docs.microsoft.com/en-us/azure/active-directory/privileged-identity-management/powershell-for-azure-ad-roles
.PARAMETER None
  None
.INPUTS
  None
.OUTPUTS
  Log file for transcription logging
.NOTES
  Version:        1.0
  Author:         Steve McIntyre     
  Creation Date:  21/05/21
  Purpose/Change: Initial Script
.LINK
  None
.EXAMPLE
  .\Register-PIMAADGroups.ps1
  Connects to Azure AD and creates Azure AD Groups and asssigned to PIM Roles
 
#>

#requires -version 4
#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  # Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

# Import Modules & Snap-ins

# Initialize your variables
#Set-Variable 

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Customer Specific

# Script scopped variables
$Script:Date                = Get-Date -Format yyyy-MM-dd                                               # Date format in yyyymmdd
$Script:File                = ''                                                                        # File var for Get-FilePicker Function
$Script:ScriptName          = 'Register-PIMAADGroups'                                                   # Script Name used in the Open Dialogue
$Script:LogFile             = $PSScriptRoot + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:GUID                = ''                                                                        # Script GUID
#^ Use New-Guid cmdlet to generate new script GUID for each version change of the script

#-----------------------------------------------------------[Hash Tables]-----------------------------------------------------------

#-----------------------------------------------------------[Functions]-------------------------------------------------------------

#& Start Transcriptions
Function Start-Logging{

    try {
        Stop-Transcript | Out-Null
    } catch [System.InvalidOperationException] { }                                          # jobs are running
    $ErrorActionPreference = "Continue"                                                     # Set Error Action Handling
    Get-Now                                                                                 # Get current date time
    Start-Transcript -path $Script:LogFile -IncludeInvocationHeader -Append                 # Start Transcription append if log exists
    Write-Host ''                                                                           # write Line spacer into Transcription file
    Write-Host ''                                                                           # write Line spacer into Transcription file
    Write-Host  "========================================================" 
    Write-Host  "====== $Script:Now Processing Started ========" 
    Write-Host  "========================================================" 
    Write-Host  ''
    
    Write-Host ''                                                                           # write Line spacer into Transcription file
  }
  
  #& Date time formatting for timestamped updated
  Function Get-Now{
    $Script:Now = (get-date).tostring("[dd/MM HH:mm:ss:ffff]")
  }

#-----------------------------------------------------------[Execution]------------------------------------------------------------
<#
? ---------------------------------------------------------- [NOTES:] -------------------------------------------------------------
& Best veiwed and edited with Microsoft Visual Studio Code with colorful comments extension
* Requires AzureAD / AzureADPreview
? Assigning Groups to PIM roles is still currently in preview - Use in production at own risk ;)
? Currently only AzureAD groups are supported you cannot assign to a group authored on Prem, nor nest OnPrem group into the AAD Group
^ $group = New-AzureADMSGroup -DisplayName "Contoso_Helpdesk_Administrators" -Description "This group is assigned to Helpdesk Administrator built-in role in Azure AD." -MailEnabled $true -SecurityEnabled $true -MailNickName "contosohelpdeskadministrators" -IsAssignableToRole $true
^ $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Helpdesk Administrator'" 
^ $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
* Transcription logging formatting use Get-Now before write-host to return current timestamp into $Scipt:Now variable
  Write-Host "$Script:Now [INFORMATION] Information Message"
  Write-Host "$Script:Now [WARNING] Warning Message"
  Write-Host "$Script:Now [ERROR] Error Message"
? ---------------------------------------------------------------------------------------------------------------------------------
#>

# Script Execution goes here

Start-Logging                                                                                       # Start Transcription logging
Get-Now                                                                                             # Get Timestamp
Write-Host "$Script:Now [INFORMATION] Connecting to Azure AD"  
Connect-AzureAD                                                                                     # Connect to AzureAD

#! Setup PIM Groups and assign to Azure AD Role
Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Helpdesk Administrators role group (PIM_Helpdesk_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Helpdesk_Admins" -Description "This group is assigned to Helpdesk Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimhelpdeskadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Helpdesk Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Global Administrator role group (PIM_Global_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Global_Admins" -Description "This group is assigned to Global Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimglobaladmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Global Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Global Reader role group (PIM_Global_Reader)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Global_Reader" -Description "This group is assigned to Global Reader built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimglobalreader" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Global Reader'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM AAD Local Administrator role group (PIM_AAD_Device_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_AAD_Device_Admins" -Description "This group is assigned to Azure AD Joined Device Local Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimaaddeviceadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Azure AD Joined Device Local Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Cloud Device Administrator role group (PIM_Cloud_Device_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Cloud_Device_Admins" -Description "This group is assigned to Cloud Device Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimclouddeviceadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Cloud Device Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Application Administrator role group (PIM_Application_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Application_Admins" -Description "This group is assigned to Application Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimclouddeviceadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Application Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Compliance Administrator role group (PIM_Compliance_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Compliance_Admins" -Description "This group is assigned to Compliance Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimcomplianceadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Compliance Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Exchange Administrator role group (PIM_Exchange_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Exchange_Admins" -Description "This group is assigned to Exchange Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimexchangeadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Exchange Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Groups Administrator role group (PIM_Groups_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Groups_Admins" -Description "This group is assigned to Groups Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimgroupsadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Groups Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Guest Inviter role group (PIM_Guest_Inviter)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Guest_Inviter" -Description "This group is assigned to Guest Inviter built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimguestinviter" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Guest Inviter'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Intune Administrator role group (PIM_Intune(MEM)_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Intune(MEM)_Admins" -Description "This group is assigned to Intune Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimintunememadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Intune Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM License Administrator role group (PIM_License_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_License_Admins" -Description "This group is assigned to License Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimlicenseadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'License Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Priviledged Auth Admin role group (PIM_PA_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_PA_Admins" -Description "This group is assigned to Privileged Authentication Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimpaadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Privileged Authentication Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM PIM Administrator role group (PIM_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Admins" -Description "This group is assigned to Privileged Role Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Privileged Role Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Security Administrator role group (PIM_Security_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Security_Admins" -Description "This group is assigned to Security Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimsecurityadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Security Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Security Operator role group (PIM_Security_Ops)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Security_Ops" -Description "This group is assigned to Security Operator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimsecops" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Security Operator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Security Reader role group (PIM_Security_Reader)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Security_Reader" -Description "This group is assigned to Security Reader built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimsecurityreader" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Security Reader'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM SharePoint Administrator role group (PIM_SharePoint_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_SharePoint_Admins" -Description "This group is assigned to SharePoint Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimsharepointadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'SharePoint Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM Teams Administrator role group (PIM_Teams_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_Teams_Admins" -Description "This group is assigned to Teams Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimteamsadmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'Teams Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes  
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Try{
    Get-Now                                                                                             # Get Timestamp
    Write-Host "$Script:Now [INFORMATION] Creating PIM User Administrators role group (PIM_User_Admins)"  
    $group = New-AzureADMSGroup -DisplayName "PIM_User_Admins" -Description "This group is assigned to User Administrator built-in role in Azure AD." -SecurityEnabled $true -MailEnabled $false -MailNickName "pimuseradmins" -IsAssignableToRole $true
    $roleDefinition = Get-AzureADMSRoleDefinition -Filter "displayName eq 'User Administrator'" 
    $roleAssignment = New-AzureADMSRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $group.Id 
    $null = $group; $null = $roleDefinition; $null = $roleAssignment                                    # Reset Variabes  
}
Catch{
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an issue creating the Group or Role assignement"
    Write-Host $PSItem.Exception.Message
}
Finally{
    $Error.Clear()                                                                                      # Clear error log
}

Get-Now
Write-Host  "========================================================" 
Write-Host  "======== $Script:Now Processing Finished =========" 
Write-Host  "========================================================"

Stop-Transcript                                                                                     # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------