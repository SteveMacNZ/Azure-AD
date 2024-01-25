<#
.SYNOPSIS
  Exports OneDrive specific information to CSV file
.DESCRIPTION
  Exports OneDrive specific information to CSV file
.PARAMETER None
  None
.INPUTS
  None
.OUTPUTS
  Log file for transcription logging
  CSV AD information required for OneDrive Migrations
.NOTES
  Version:        1.1
  Author:         Steve McIntyre
  Creation Date:  17/06/19
  Purpose/Change: Updated with FilePicker and exporting onedrives with usage + quotas
.EXAMPLE
  .\OD-ExportAD.ps1
  
#>

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
$CustName = Read-Host "Enter Customer Short Name (e.g. Cloud Innovation = CINZ)"            # Customer Short Name
$ADBase = Read-Host "Enter the AD SearchBase string (e.g. DC=domain,DC=co,DC=nz)"           # AD Search Base

# Script scopped variables
$Script:Date = Get-Date -Format FileDate                                                    # Date format in yyyymmdd
$Script:LogFile = $PSScriptRoot + "\" + $Script:Date + "_" + $CustName + "_OD-ExportAD.log" # logfile location and name
$Script:CSVFile = $PSScriptRoot + "\" + $Script:Date + "_" + $CustName + "_AD.csv"          # CSV Export location and name

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

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Script Execution goes here

Start-Logging                                                                               # Call Start-Logging function to start transcription

Write-Output "Querying $ADBase for OneDrive Specific Information"                           # Write Status Update to Transcription file

# Query AD
Get-ADUser -searchbase $ADBase -Filter * `
-Properties distinguishedname,name,samaccountname,homedrive,homedirectory,mail,userprincipalname,department,office,whencreated,`
lastlogondate,enabled | Select-Object name,samaccountname,homedrive,homedirectory,mail,userprincipalname,`
department,office,distinguishedname,whencreated,lastlogondate,enabled | Export-Csv -Path $Script:CSVFile

Write-Output ''                                                                             # write Line spacer into Transcription file
Write-Output "AD Information has been written out to $Script:CSVFile"                       # Write Status Update to Transcription file

Stop-Transcript                                                                             # Stop Transcription