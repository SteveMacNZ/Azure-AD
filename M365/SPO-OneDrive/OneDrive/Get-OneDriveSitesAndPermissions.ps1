<#
.SYNOPSIS
  Connects to SharePoint Online and enumerates all user personal OneDrive sites and permissions
.DESCRIPTION
  Connects to SharePoint Online and enumerates all user personal OneDrive sites and permissions 
.PARAMETER None
  None
.INPUTS
  What Inputs  
.OUTPUTS
  What outputs
.NOTES
  Version:        1.0.0.0
  Author:         Steve McIntyre
  Creation Date:  DD/MM/20YY
  Purpose/Change: Initial Release
.LINK
  None
.EXAMPLE
  ^ . Get-OneDriveSitesAndPermissions.ps1
  does what with example of cmdlet
  Invoke-What.ps1

#>

#requires -version 4
#region ------------------------------------------------------[Script Parameters]--------------------------------------------------

Param (
  #Script parameters go here
)

#endregion
#region ------------------------------------------------------[Initialisations]----------------------------------------------------

#& Global Error Action
#$ErrorActionPreference = 'SilentlyContinue'

#& Module Imports
#Import-Module ActiveDirectory

#& Includes - Scripts & Modules
. Get-CommonFunctions.ps1                                                           # Include Common Functions

#endregion
#region -------------------------------------------------------[Declarations]------------------------------------------------------

# Script sourced variables for General settings and Registry Operations
$Script:Date        = Get-Date -Format yyyy-MM-dd                                   # Date format in yyyy-mm-dd
$Script:Now         = ''                                                            # script sourced veriable for Get-Now function
$Script:ScriptName  = 'Get-OneDriveSitesAndPermissions'                             # Script Name used in the Open Dialogue
$Script:dest        = "$PSScriptRoot\Exports"                                       # Destination path
$Script:LogDir      = "$PSScriptRoot\Logs"                                          # Logdir for Clear-TransLogs function for $PSScript Root
$Script:LogFile     = $Script:LogDir + "\" + $Script:Date + "_" + $env:USERNAME + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:CSVFile     = $Script:dest + "\" + $Script:Date + "_ODSites&Perms.csv"      # CSV Export location and name
$Script:BatchName   = ''                                                            # Batch name variable placeholder
$Script:GUID        = '69237689-f216-4c93-a54b-f4ed2717b723'                        # Script GUID
  #^ Use New-Guid cmdlet to generate new script GUID for each version change of the script
[version]$Script:Version  = '1.0.0.0'                                               # Script Version Number
$Script:Client      = 'Meridian Energy Limited'                                     # Set Client Name - Used in Registry Operations
$Script:WHO         = whoami                                                        # Collect WhoAmI
$Script:Desc        = "One Drive Sites and Permissions Report"                      # Description displayed in Get-ScriptInfo function
$Script:Desc2       = "Collects information on User OneDrive Sites and Permissions and exports to CSV" # Description2 displayed in Get-ScriptInfo function
$Script:PSArchitecture = ''                                                         # Place holder for x86 / x64 bit detection

#$Script:TenantName  = "meridianenergy"                                              # Tenant name for connecting to sharepoint admin site
$Script:TenantName  = "m365x52636157"
#^ Array lists
$Script:ODResults   = [System.Collections.ArrayList]@()                             # Arraylist of OneDrive Results from Class

#endregion
#region --------------------------------------------------------[Hash Tables]------------------------------------------------------

#& any script specific hash tables that are not included in Get-CommonFunctions.ps1

#endregion
#region -------------------------------------------------------[Functions]---------------------------------------------------------

#& any script specific funcitons that are not included in Get-CommonFunctions.ps1

#endregion
#region ------------------------------------------------------------[Classes]-------------------------------------------------------------

#& any script specific classes that are not included in Get-CommonFunctions.ps1

# Example Class - constuct and usage
Class SPO_OneDrive{
  # $classresult = [ClassName]::new("$WhatString","$WhatINT","$WhatBool")           # creates a new class object
  # $Script:ClassArray.add($classresult) | Out-Null                                 # writes the class object to the Class array
  # $Script:ClassArray | Export-Csv -Path $ClassReport -NoTypeInformation           # writes the class array out to CSV file
  [String]$Name
  [String]$Owner
  [String]$Admins
  [String]$Permissions
  [String]$URL
  [String]$Status
  [String]$LockState
  [String]$LastContentModifiedDate
  [INT]$UsedMB 
  [INT]$UsedGB
  [INT]$StorageQuotaGB
  [INT]$StorageQuotaWarnGB
  [INT]$ResourceQuotaGB
  [INT]$ResourceQuotaWarnGB
    
  # constructor
  SPO_OneDrive([String]$Name, [String]$Owner, [String]$Admins, [String]$Permissions, [String]$URL, [String]$Status, [String]$LockState, [String]$LastContentModifiedDate, [INT]$UsedMB, [INT]$UsedGB, [INT]$StorageQuotaGB, [INT]$StorageQuotaWarnGB, [INT]$ResourceQuotaGB, [INT]$ResourceQuotaWarnGB){
    $this.Name = $Name
    $this.Owner = $Owner
    $this.Admins = $Admins
    $this.Permissions = $Permissions
    $this.URL = $URL
    $this.Status = $Status
    $this.LockState = $LockState
    $this.LastContentModifiedDate = $LastContentModifiedDate
    $this.UsedMB = $UsedMB
    $this.UsedMB = $UsedMB
    $this.StorageQuotaGB = $StorageQuotaGB
    $this.StorageQuotaWarnGB = $StorageQuotaWarnGB
    $this.ResourceQuotaGB = $ResourceQuotaGB
    $this.ResourceQuotaWarnGB = $ResourceQuotaWarnGB    
  } 
}

#endregion
#region -----------------------------------------------------------[Execution]------------------------------------------------------------
<#
? ---------------------------------------------------------- [NOTES:] -------------------------------------------------------------
& Best veiwed and edited with Microsoft Visual Studio Code with colorful comments extension
^ Transcription logging formatting use the following functions to Write-Host messages
  Write-InfoMsg "Message" writes informational message as Write-Host "$Script:Now [INFORMATION] Information Message" format
  Write-InfoHighlightedMsg "Message" writes highlighted information message as Write-Host "$Script:Now [INFORMATION] Highlighted Information Message" format
  Write-SuccessMsg "Message" writes success message as Write-Host "$Script:Now [SUCCESS] Warning Message" format"
  Write-WarningMsg "Message" writes warning message as Write-Host "$Script:Now [WARNING] Warning Message" format
  Write-ErrorMsg "Message" writes error message as Write-Host "$Script:Now [ERROR] Error Message" format
  Write-ErrorAndExitMsg "Message" writes error message as Write-Host "$Script:Now [ERROR] Error Message" format and exits script
? ---------------------------------------------------------------------------------------------------------------------------------
#>

Start-Logging                                                                           # Start Transcription logging
Get-PSArch                                                                              # Get PS Architecture
Get-ScriptInfo                                                                          # Display Script Info
Clear-TransLogs                                                                         # Clear logs over 15 days old

Invoke-TestPath -ParamPath $Script:dest                                                 # Test and create folder structure 
Invoke-TestPath -ParamPath $Script:LogDir                                               # Test and create folder structure

$TenantUrl = "https://$Script:TenantName-admin.sharepoint.com"                          # Create SharePoint Admin URL based on $TenantName
Write-Output "Connecting to $TenantURL now"                                             # Write Status Update to Transcription file

Connect-SPOService -Url $TenantUrl                                                      # Connect to SharePoint PowerShell Session using modern auth

$OneDriveSites = Get-SPOSite -Template "SPSPERS" -Limit ALL -includepersonalsite $True  # Get all Personal Site collections

$counter = 0                                                                            # Init loop counter
$maximum = $OneDriveSites.Count                                                         # number of items to be processed

Write-InfoHighlightedMsg "$maximum User Personal OneDrive Sites found"
Write-Host ""
Foreach ($site in $OneDriveSites)  {
  
  $counter++
  $percentCompleted = $counter * 100 / $maximum

  $message = '{0:p1} completed, processing {1}.' -f ( $percentCompleted/100), $site.Title
  Write-Progress -Activity 'I am busy' -Status $message -PercentComplete $percentCompleted

  Write-InfoMsg "Processing OneDrive for $($site.Title)"

  $odName = $Site.Title                                                                 # OneDrive Site Title = User Name
  $odOwner = $Site.Owner                                                                # OneDrive owner in UPN formation
  $odURL = $Site.Url                                                                    # OneDrive URL
  $odStatus = $Site.Status                                                              # OneDrive Status
  $odLockState = $Site.LockState                                                        # OneDrive LockState
  $odLastContentModifiedDate = $Site.LastContentModifiedDate                            # OneDrive Date last modified
  $odUsedMB = $Site.StorageUsageCurrent                                                 # OneDrive current used storage in MB
  $odUsedGB = $Site.StorageUsageCurrent/1024                                            # OneDrive current used storage in GB
  $odStorageQuotaGB = $Site.StorageQuota/1024                                           # OneDrive Storage quota in GB
  $odStorageQuotaWarnGB = $Site.StorageQuotaWarningLevel/1024                           # OneDrive Storage Quota Warning level in GB
  $odResourceQuotaGB = $Site.ResourceQuota/1024                                         # OneDrive Resource Quota in GB
  $odResourceQuotaWarnGB = $Site.ResourceQuotaWarningLevel/1024                         # OneDrive Resource Quota Warning level in GB

  Try{
    Write-InfoMsg "Collecting admin permissions on $($Site.Title) OneDrive"
    $Admins = Get-SPOUser -Site $odURL | Where-Object {$_.IsSiteAdmin -eq $true}        # Get all users who have admin access on users ondrive
    Foreach ($Admin in $Admins){
      $odAdmins += $Admin.LoginName + "; "                                              # Add admin upn to admins list  
    }
    $odAdmins = $odAdmins.TrimEnd("; ")                                                 # trim ending ; from admin list

    Write-InfoMsg "Collecting user permissions on $($Site.Title) OneDrive"
    $Perms = Get-SPOUser -Site $odURL | Where-Object {$_.IsSiteAdmin -eq $false}        # Get all users who have access on users ondrive
    Foreach ($Perm in $Perms){
      $odPerms += $Perm.LoginName + "; "                                                # Add admin upn to users list  
    }
    $odPerms = $odPerms.TrimEnd("; ")                                                   # trim ending ; from users list
  }
  Catch{
    Write-ErrorMsg "Unable to collect permissions on $($Site.Title) OneDrive"
    Write-Host $PSItem.Exception.Message -ForegroundColor RED                           # Error message details
  }
  Finally{
    $Error.Clear()                                                                      # Clear error log
  }

  # Create and populate SharePoint OneDrive class object
  $SPO_OD_result = [SPO_OneDrive]::new("$odName","$odOwner","$odAdmins","$odPerms","$odURL","$odStatus","$odLockState","$odLastContentModifiedDate","$odUsedMB","$odUsedGB","$odStorageQuotaGB","$odStorageQuotaWarnGB","$odResourceQuotaGB","$odResourceQuotaWarnGB")
  Write-SuccessMsg "$($Site.Title) written to class object"
  $Script:ODResults.add($SPO_OD_result) | Out-Null

  # Clear all defined variables
  ($odName,$odOwner,$odURL,$odStatus,$odLockState,$odLastContentModifiedDate,$odUsedMB) = $null
  ($odUsedGB, $odStorageQuotaGB,$odStorageQuotaWarnGB,$odResourceQuotaGB) = $null
  ($odResourceQuotaWarnGB,$odAdmins,$odPerms,$SPO_OD_result,$Admins) = $null

  Write-Host ""
}


Write-InfoMsg "Writing class objects to csv file"
$Script:ODResults | Export-Csv -Path $Script:CSVFile -NoTypeInformation

# Input / Output comparsion
Write-Host ""
Write-Host '--------------------------------------------------------------------------------'
Write-Host '|                      Input / Output CSV Count Comparsion                     |'
Write-Host '--------------------------------------------------------------------------------'
$OutputObject = Import-csv $Script:CSVFile -Delimiter ","        # Read Output for input/output comparsion
$OutputCount = $OutputObject.count
If ($maximum -eq $OutputCount){
  Get-Now
  Write-Host "$Script:Now [COUNTS] CSV Input [$maximum] and Output [$OutputCount] counts match" @chighlight
}
else{
  Get-Now
  Write-Host "$Script:Now [COUNTS] CSV Input [$maximum] and Output [$OutputCount] counts don't match" @cerror
}
Write-Host '--------------------------------------------------------------------------------'
Write-Host ''
#>

Write-Host ''
Get-Now
Write-Host "$Script:Now [INFORMATION] Processing finished + any outputs"                          

Get-Now
Write-Host "================================================================================"  
Write-Host "================= $Script:Now Processing Finished ====================" 
Write-Host "================================================================================" 

Stop-Transcript
#endregion
#---------------------------------------------------------[Execution Completed]----------------------------------------------------------