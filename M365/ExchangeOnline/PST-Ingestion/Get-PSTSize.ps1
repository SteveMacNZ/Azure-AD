<#
.SYNOPSIS
  Get PST Names And Sizes from a directory
.DESCRIPTION
  Get PST Names And Sizes from a directory - Assumes all PST files in a single directory
.PARAMETER None
  None
.INPUTS
  Source Folder selected containing PST files
.OUTPUTS
  Log file stored in same location as the CSV file
  CSV output of 
.NOTES
  Version:        1.1
  Author:         Steve McIntyre
  Creation Date:  18/08/2021
  Purpose/Change: Updated to new template with new functions added for logging and reporting
  Version:        1.0
  Author:         Steve McIntyre
  Creation Date:  12/10/19
  Purpose/Change: Initial Script
.LINK
  https://github.com/SteveMacNZ/Azure-M365/tree/main/M365/ExchangeOnline/PST-Ingestion
.EXAMPLE
  .\Get-PSTSize.ps1
  Gets all PST files in specified directory and returns their size
#>

#requires -version 4
#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  #Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

# Import Modules & Snap-ins

# Initialize your variables
#Set-Variable 

#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Customer Specific
#$CustName = Read-Host "Enter Customer Short Name (e.g. Cloud Innovation = CINZ)"                        # Customer Short Name

# Script scopped variables
$Script:Date                = Get-Date -Format yyyy-MM-dd                                               # Date format in yyyymmdd
$Script:File                = ''                                                                        # File var for Get-FilePicker Function
$Script:ScriptName          = 'Get-PSTSize'                                                             # Script Name used in the Open Dialogue
$Script:LogFile             = $PSScriptRoot + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:CSVFile             = $PSScriptRoot + "\" + $Script:Date + "_" +"_PSTSizes.csv"                 # CSV Export location and name
$Script:GUID                = '127543ce-4388-4bd8-a97c-beef1ca4b8d5'                                    # Script GUID
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

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Script Execution goes here
Start-Logging                                                                                       # Start Transcription logging
Clear-TransLogs                                                                                     # Clear logs over 15 days old

Get-Now
Write-Host "$Script:Now [INFORMATION] Script processing started"

$PSTPath = Read-Host "Enter the PST Path (e.g. C:\Temp\<CUST>\PST)"                                 # Path to PST Files
Get-Now                                                                                             # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] $PSTPath has been selected for processing" -ForegroundColor Magenta
Write-Host ""

$Result = Get-ChildItem -Path $PSTPath | Select-Object Name, @{Name="GB";Expression={ "{0:N2}" -f ($_.Length / 1GB) }}  # Get PST files and length converted to GB
$Result | Format-Table                                                                              # Format results as a table
$Result | Select-Object Name,GB | Export-Csv $Script:CSVFile -NoTypeInformation                     # Export the data to CSV

Write-Host ''                                                                                       # write Line spacer into Transcription file
Get-Now
Write-Host "$Script:Now [INFORMATION] Processing finished, details written to $Script:CSVFile"      # Write Status Update to Transcription file

Get-Now
Write-Host  "========================================================" 
Write-Host  "======== $Script:Now Processing Finished =========" 
Write-Host  "========================================================"

Stop-Transcript                                                                                     # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------