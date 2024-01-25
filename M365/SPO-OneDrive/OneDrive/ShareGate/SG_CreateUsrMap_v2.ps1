<#
.SYNOPSIS
  ShareGate: Desktop - Create User Mapping .sgum file
.DESCRIPTION
  ShareGate: Desktop - Create User Mapping .sgum file
.PARAMETER None
  None
.INPUTS
  CSV File with user mapping containing the following
  samaccount name               One SamAccountName per row
  owner                         Owner in format of Email Address  / UPN
.OUTPUTS
  Log file for transcription logging
  .sgum file containing the user mapping
.NOTES
  Version:        1.0
  Author:         Steve McIntyre
  Creation Date:  17/06/19
  Purpose/Change: Updated with FilePicker
.EXAMPLE
  .\SG_CreateUsrMap_v2.ps1
  
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  # Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

# Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Customer specific
$CustName = Read-Host "Enter the customer name (e.g. MOE or MeridianEnergy)"            # Name of the customer  

# Script scopped variables
$Script:File = ''                                                                       # File
$Script:Date = Get-Date -Format FileDate                                                # Date format in yyyymmdd
$Script:LogFile = $PSScriptRoot + "\" + $Script:Date + "_SG_UserMapping.log"            # logfile location and name
$Script:MapPath = $PSScriptRoot + "\" + $Script:Date + "_" + $CustName + ".sgum"        # CSV Export location and name

$mappingSettings = New-MappingSettings                                                  # Declaration of the mapping settings:

#-----------------------------------------------------------[Functions]------------------------------------------------------------

# Start Transcriptions
Function Start-Logging{
    # Create Log File + Start Logging
    if ($Null -ne $Script:File) {
        # $Log = $Script:File + ".log"
        $ErrorActionPreference="SilentlyContinue"
        Stop-Transcript | out-null
        $ErrorActionPreference = "Continue"
        Start-Transcript -path $Script:LogFile -append
    }
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

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Script Execution goes here

Get-FilePicker                                                                          # Call Get-FilePicker Function

Start-Logging                                                                           # Call Start-Logging function to start transcription

Write-Output "importing $Script:File for processing"                                    # Write status update to Transcript
$table = Import-CSV $Script:File -Delimiter ","                                         # Create a table based on the csv

# Cycle through each row of the CSV
foreach ($row in $table) {

    # Add the current row source user and destination user to the mapping list
    Write-Output "Setting Mapping Settings for $row.samaccountname"                     # Write status update to Transcript
    Write-Output ''
    Set-UserAndGroupMapping -MappingSettings $mappingSettings -Source $row.samaccountname -Destination $row.Owner

}

# The user and group mappings are exported to scriptroot
Write-Output "Exporting user map to file"                                               # Write status update to Transcript
Write-Output ''
Export-UserAndGroupMapping -MappingSettings $mappingSettings -Path $Script:MapPath
Write-Output "User mapping written to $Script:MapPath"                                  # Write status update to Transcript
Write-Output ''

Stop-Transcript                                                                         # Stop Transcription logging