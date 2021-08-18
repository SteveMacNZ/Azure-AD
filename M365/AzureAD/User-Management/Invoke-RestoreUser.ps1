<#
.SYNOPSIS
  Restores a deleted users in O365 portal from CSV file and removes the ImmutableID, uses legacy MsolService for restore of User and AzureAD module for remaining steps
.DESCRIPTION
  Restores a deleted users in O365 portal from CSV file and removes the ImmutableID, uses legacy MsolService for restore of User and AzureAD module for remaining steps,
  currently the Restore-AzureADMSDeletedDirectoryObject only supports groups and applications for restoring. To use native AzureAD commands only comment out the MsolService
  commands and follow https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/users-bulk-restore to restore uses in bulk prior to running this script
.PARAMETER None
  None
.INPUTS
  CSV file with following fields (Use CloudUsers.xlsx template)
  UserPrincipalName Email addresss / UPN of the User
.OUTPUTS
  Transcription log stored in the script root directory
.NOTES
  Version:        1.3
  Author:         Steve McIntyre
  Creation Date:  18/08/2021
  Purpose/Change: Updates for function changes and logic for inital upload to GitHub + Updated to use AzureAD module
  Version:        1.1
  Author:         Steve McIntyre
  Creation Date:  01/10/19
  Purpose/Change: Standardisation Updates
.LINK
  https://github.com/SteveMacNZ/Azure-M365/tree/main/M365/Azure-AD/User-Management
.EXAMPLE
  .\Invoke-RestoreUser.ps1
  Invokes restore of user accounts using AzureAD
#>

#requires -version 4 -Modules AzureAD, MsolService
#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  #Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Customer Specific
#$CustName = Read-Host "Enter Customer Short Name (e.g. Cloud Innovation = CINZ)"                        # Customer Short Name

# Script scopped variables
$Script:Date                = Get-Date -Format yyyy-MM-dd                                               # Date format in yyyymmdd
$Script:File                = ''                                                                        # File var for Get-FilePicker Function
$Script:ScriptName          = 'Invoke-RestoreUser'                                                      # Script Name used in the Open Dialogue
$Script:LogFile             = $PSScriptRoot + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:GUID                = '15726e2d-4da7-451f-9e74-c458de26ac8a'                                    # Script GUID
#^ Use New-Guid cmdlet to generate new script GUID for each version change of the script

#-----------------------------------------------------------[Functions]------------------------------------------------------------
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

#& Clean up log files in script root older than 15 days
Function Clear-TransLogs{
  Get-Now
  Write-Output "$Script:Now - Cleaning up transaction logs over 15 days old"
  Get-ChildItem $PSScriptRoot -recurse "*$Script:ScriptName.log" -force | Where-Object {$_.lastwritetime -lt (get-date).adddays(-15)} | Remove-Item -force
}

#& FilePicker function for selecting input file via explorer window
Function Get-FilePicker {
  Param ()
  [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
  $ofd = New-Object System.Windows.Forms.OpenFileDialog
  $ofd.InitialDirectory = $PSScriptRoot                                                         # Sets initial directory to script root
  $ofd.Title            = "Select file for $Script:ScriptName"                                  # Title for the Open Dialogue
  $ofd.Filter           = "Text files (*.txt)|*.txt|CSV File (*.csv)|*.csv|All files (*.*)|*.*" # Display All files / Txt / CSV
  $ofd.FilterIndex      = 2                                                                     # 3 Default to display All files
  $ofd.RestoreDirectory = $true                                                                 # Reset the directory path
  #$ofd.ShowHelp         = $true                                                                 # Legacy UI              
  $ofd.ShowHelp         = $false                                                                # Modern UI
  if($ofd.ShowDialog() -eq "OK") { $ofd.FileName }
  $Script:File = $ofd.Filename
}

#& Test-AAD function tests for exisitng connection to Azure AD and if session does not exists connects 
Function Test-AAD {
  try {
    #^ Attempts to return Tenant DisplatName to verify connection to Azure AD
    $IsAADConnected = Get-AzureADTenantDetail | Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue
  }
  Catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]{}  
  
  if ($null -eq $IsAADConnected){
    Get-Now
    Write-Host "$Script:Now [INFORMATION] Azure AD - is not connected - Connecting...." -ForegroundColor Magenta
    Connect-AzureAD                                                                         # Connect to Azure AD using Modern Auth
  }
  else {
    Get-Now
    write-host "$Script:Now [INFORMATION] Already connected to AzureAD - Proceeding" -ForegroundColor Green 
  }
}

#& Test-Msol function tests for exisitng connection to MsolService and if session does not exists connects 
Function Test-Msol {
  try {
    #^ Attempts to return Tenant Domain Name to verify connection to MsolService
    $IsMsolConnected = Get-MsolDomain -ErrorAction SilentlyContinue
  }
  Catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException]{}  
  
  if ($null -eq $IsMsolConnected){
    Get-Now
    Write-Host "$Script:Now [INFORMATION] MsolService - is not connected - Connecting...." -ForegroundColor Magenta
    Connect-MsolService                                                                         # Connect to MsolService
  }
  else {
    Get-Now
    write-host "$Script:Now [INFORMATION] Already connected to MsolService - Proceeding" -ForegroundColor Green 
  }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
# Script Execution goes here

Start-Logging                                                                                       # Start Transcription logging
Clear-TransLogs                                                                                     # Clear logs over 15 days old

Get-Now
Write-Host "$Script:Now [INFORMATION] Script processing started"

Write-Host ""
Get-FilePicker                                                                                      # Call FilePicker Function
Get-Now                                                                                             # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] $Script:File has been selected for processing" -ForegroundColor Magenta
Write-Host ""

Test-AAD                                                                                            # Test for connection to Azure AD
Test-Msol                                                                                           # Test for connection to MsolService

$Restores = Import-csv $Script:File -Delimiter ","                                                  # Load CSV into array
$tc = $Restores.count                                                                               # Count number of imports to be completed
$lc = 0

ForEach ($Restore in $Restores) {
  # set up progress notification
  $lc++
  Write-Progress "Processing $tc Objects" -Status "Completed: $lc of $tc. Remaining: $($tc-$lc)" -PercentComplete ($lc/$tc*100)
  
  # Try restoring user account
  Try {
    Get-Now
    Write-Host "$Script:Now [INFORMATION] Restoring "$_.UserPrincipalName"" -ForegroundColor Yellow # Write Status Update to Transaction Log
    Restore-MsolUser -UserPrincipalName $_.UserPrincipalName                                        # Restore Deleted User account
  }
  Catch {
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an error Restoring account for "$_.UserPrincipalName"" -ForegroundColor Red
    Write-Host $PSItem.Exception.Message -ForegroundColor RED
  }
  Finally{
    $Error.Clear()                                                                                  # Clear error log
  }

  # Try resetting password
  Try {
    Get-Now
    Write-Host "$Script:Now [INFORMATION] Resetting "$_.UserPrincipalName" Password"                # Write Status Update to Transaction Log
    Set-AzureADUserPassword -ObjectId $_.UserPrincipalName -EnforceChangePasswordPolicy $True       # Reset Users password and force change on login
  }
  Catch {
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an error resetting password for "$_.UserPrincipalName"" -ForegroundColor Red
    Write-Host $PSItem.Exception.Message -ForegroundColor RED
  }
  Finally{
    $Error.Clear()                                                                                  # Clear error log
  }

  # Try removing ImmutableID
  Try {
    Get-Now
    Write-Host "$Script:Now [INFORMATION] Removing ImmutableID from "$_.UserPrincipalName""         # Write Status Update to Transaction Log
    Set-AzureADUser -ObjectId $_.UserPrincipalName -ImmutableID "$Null"                             # Removes ImmutableID from User account where they were a hybrid account
  }
  Catch {
    Get-Now
    Write-Host "$Script:Now [ERROR] There was an error removing the ImmutableID for "$_.UserPrincipalName"" -ForegroundColor Red
    Write-Host $PSItem.Exception.Message -ForegroundColor RED
  }
  Finally{
    $Error.Clear()                                                                                  # Clear error log
  }

}

Get-Now                                                                                             # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] Restoration of User accounts has been completed, Please see Transcription log for temporary passwords!"

Get-Now
Write-Host  "========================================================" 
Write-Host  "======== $Script:Now Processing Finished =========" 
Write-Host  "========================================================"

Stop-Transcript                                                                                     # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------