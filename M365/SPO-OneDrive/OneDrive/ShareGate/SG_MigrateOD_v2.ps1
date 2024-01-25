<#
.SYNOPSIS
  ShareGate: Desktop - Migrate Fileshare to OneDrive
.DESCRIPTION
  ShareGate: Desktop - Migrate Fileshare to OneDrive
.PARAMETER None
  None
.INPUTS
  CSV File containing;
    Url                   OneDrive Personal Url
    Owner                 User Email address
    homedirectory         UNC path to HomeShare directory

 .sgum file               Containing the user mapping
 .sgt file                Containind the file exclusions
.OUTPUTS
  Log file for transcription logging
  CSV files for each migration loop stored in $Script:RepDir
.NOTES
  Version:        1.1
  Author:         Steve McIntyre
  Creation Date:  17/06/19
  Purpose/Change: Updated with FilePicker
.EXAMPLE
  .\SG_MigrateOD_v2.ps1
  
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  # Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

# Import Modules & Snap-ins

# Initialize your variables
Set-Variable dstSite, dstList

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Script scopped variables
$Script:File = ''                                                                       # File
$Script:Date = Get-Date -Format FileDate                                                # Date format in yyyymmdd
$Script:LogFile = $PSScriptRoot + "\" + $Script:Date + "_SG_DataMigration.log"          # logfile location and name
$Script:RepDir = "$PSScriptRoot\Reports"                                                # Reports directory

$usermap = "$PSScriptRoot\yyymmdd_CST.sgum"                                             # Path to the User Mapping file
$excludefiles = "$PSScriptRoot\CST_Exc.sgt"                                             # Path to the File exclusions
$MigUsr = Read-Host "Enter the username of the account to use for Migration"            # username of the migration account
$EnterPwd = Read-Host "Enter the password for the migration user account"               # Password for the migration account
$ENCpassword = ConvertTo-SecureString $EnterPwd -AsPlainText -Force                     # Encrypted Password



#-----------------------------------------------------------[Functions]------------------------------------------------------------

# Start Transcriptions
Function Start-Logging{
  Stop-Transcript | out-null                                                                # Ensure no transcription jobs are running
  $ErrorActionPreference = "Continue"                                                       # Set Error Action Handling
  Start-Transcript -path $Script:LogFile -append                                            # Start Transcription append if log exists
  Write-Output ''                                                                           # write Line spacer into Transcription file
  Write-Output ''                                                                           # write Line spacer into Transcription file
  Write-Output '======================================Transcription started=======================================' 
  Write-Output ''                                                                           # write Line spacer into Transcription file
}

# Ask the user for input CSV File - via explorer window
Function Get-FilePicker {
    Param ()
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.ShowHelp=$true
    if($ofd.ShowDialog() -eq "OK") { $ofd.FileName }
    $Script:File = $ofd.Filename
    Write-Output "$Script:File was selected"
  }

# Test for Reports Directory and create if not exisitng
Function Test-ReportDir{
    # Test for report Folder
    if (!(Test-path -Path "$PSScriptRoot\Reports")){
        Write-Output "Reports folder does not exist creation folder now"
        New-Item -Path "$PSScriptRoot\Reports" -ItemType "directory"
        Write-Output "Reports folder has been created"
        Write-Output ''
    }
    else {
        Write-Output "Reports directory exist contuining...."
        Write-Output ''
    }
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Script Execution goes here
Start-Logging                                                                           # Call Start-Logging function to start transcription
Get-FilePicker                                                                          # Call Get-FilePicker Function

Write-Output ''                                                                         # Spacer line in Transcription

Test-ReportDir                                                                          # Testing for Reporting directory existing

$mappingSettings = Import-UserAndGroupMapping -Path $usermap                            # Import user and group mappings

Write-Output "importing $Script:File for processing"                                    # Write status update to Transcript
$table = Import-CSV $Script:File -Delimiter ","                                         # Create a table based on the csv

# Cycle through each row of the CSV
foreach ($row in $table) {

    #clear your variables to avoid any misdirection issues if one connection fails in the process
    Clear-Variable dstSite
    Clear-Variable dstList

    #connect to the destination OneDrive URL
    Write-Output ''                                                                     # Spacer line in Transcription
    Write-Output "Connecting to SharePoint Personal Site $row.URL"                      # Write status update to Transcript
    $dstSite = Connect-Site -Url $row.Url -UserName $MigUsr -Password $ENCpassword
    
    #select destination document library, name Documents by default in OneDrive
    $dstList = Get-List -Name Documents -Site $dstSite
    $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate -OnError SkipAllVersions -OnSiteObjectExists merge -OnWarning Continue
    Import-PropertyTemplate -path $excludefiles -List $dstList -Overwrite
    
    #Copy the content from your source directory to the Documents document library in OneDrive
    Write-Output "Migrating $row.Owner"
    $result = Import-Document -SourceFolder $row.homedirectory -DestinationList $dstList -MappingSettings $mappingSettings -TemplateName MOE_Exc -CopySettings $copysettings -NormalMode
    
    #Export a report for each OneDrive migration with the session ID as the title in a designated destination
    Write-Output ''
    Write-Output 'Write migration report to CSV'
    Export-Report $result -Path $Script:RepDir
    Write-Output ''
}

Write-Output "Fileshare migrations to OneDrive completed."                              # Write status update to Transcript
Write-Output "The migration reports are available at $Script:RepDir"                    # Write status update to Transcript
Write-Output ''

Stop-Transcript                                                                         # Stop Transcription logging