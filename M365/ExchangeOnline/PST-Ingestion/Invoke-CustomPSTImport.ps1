<#
.SYNOPSIS
  Import PST files from custom Azure Blob Storage
.DESCRIPTION
  Import PST files from custom Azure Blob Storage, rather than the default free blob, use AZ Copy to upload PST files
.PARAMETER None
  None
.INPUTS
  Source Azure Blob Storage + SAS URI
  CSV file with following fields 
  UserPrincipalName:    Email addresss / UPN of the User
  PSTPath:              Path of the PST file in the custom Azure Blob Storage
.OUTPUTS
  Log file stored in same location as the CSV file
  CSV output of 
.NOTES
  Version:        1.0
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
  .\Invoke-CustomPSTImport.ps1
  Invokes import of PST file from Custom Azure Blob into Exchange Online
#>

#requires -version 4 -Modules ExchangeOnlineManagement
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
$Script:ScriptName          = 'Invoke-CustomPSTImport'                                                  # Script Name used in the Open Dialogue
$Script:LogFile             = $PSScriptRoot + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:BatchName           = ''                                                                        # placeholder for batch name csv filename used
$Script:SASToken            = ''                                                                        # SAS URI Token here
$Script:TargetFolder        = '/'                                                                       # Import into root of mailbox
$Script:BadItemsLimit       = '10'                                                                      # Bad Item limit count
$Script:LargeItemsLimit     = '10'                                                                      # Large Item limit count       
$Script:GUID                = 'a2addf6c-a819-4631-a132-60d289ad484a'                                    # Script GUID
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

#& Test-ExchangeOnline function to detect connection to exchange online 
Function Test-ExchangeOnline
{
    #Check for current open O365 sessions and allow the admin to either use the existing session or create a new one
    $EXOsession = Get-PSSession | Where-Object { ($_.ComputerName -eq 'Outlook.office365.com') -and ($_.ConfigurationName -eq 'Microsoft.Exchange') }
    if($null -ne $EXOsession)
    {
        $a = Read-Host "An open session to Exchange Online PowerShell already exists. Do you want to use this session?  Enter y to use the open session, anything else to close and open a fresh session."
        if($a.ToLower() -eq 'y')
        {
            Get-Now
            Write-Host "$Script:Now [INFORMATION] Using existing Exchange Online Powershell session." -ForeGroundColor Green
            return
        }
        Disconnect-ExchangeOnline -Confirm:$false 
    }
    Import-Module ExchangeOnlineManagement
    #^ Uncomment -Prefix line to use command prefix e.g. Get-CloudMailbox
    #Connect-ExchangeOnline -Prefix "Cloud" -ShowBanner:$false
    Connect-ExchangeOnline -ShowBanner:$false
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Script Execution goes here
Start-Logging                                                                                       # Start Transcription logging
Clear-TransLogs                                                                                     # Clear logs over 15 days old

Get-Now
Write-Host "$Script:Now [INFORMATION] Script processing started"

Test-ExchangeOnline                                                                                 # Test connection to Exchange Online

Write-Host ""
Get-FilePicker                                                                                      # Call FilePicker Function
Get-Now                                                                                             # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] $Script:File has been selected for processing" -ForegroundColor Magenta
Write-Host ""

#* Sets Batch Name to be name of the file selected in FilePicker function
$BatchNameTemp      = $Script:File.split("\")[-1]
$Script:BatchName   = $BatchNameTemp.substring(0,($BatchNameTemp.length-4))

$PSTImports = Import-csv $Script:File -Delimiter ","                                                # Load CSV into array
$tc = $PSTImports.count                                                                             # Count number of imports to be completed
$lc = 0

ForEach ($PST in $PSTImports) {
    # set up progress notification
    $lc++
    Write-Progress "Processing $tc Objects" -Status "Completed: $lc of $tc. Remaining: $($tc-$lc)" -PercentComplete ($lc/$tc*100)  
    Try {
        Get-Now
        Write-Host "$Script:Now [INFORMATION] Processing "$_.UserPrincipalName"" -ForegroundColor Yellow
        New-MailboxImportRequest -Name $Script:BatchName -Mailbox $_."UserPrincipalName" -AzureBlobStorageAccountUri $_."PSTPath" `
        -AzureSharedAccessSignatureToken $Script:SASToken -TargetRootFolder $Script:TargetFolder -BadItemLimit $Script:BadItemsLimit -LargeItemLimit $Script:LargeItemsLimit
    }
    Catch {
        Get-Now
        Write-Host "$Script:Now [ERROR] There was an error with the PST import for "$_.UserPrincipalName"" -ForegroundColor Red
        Write-Host $PSItem.Exception.Message -ForegroundColor RED
    }
    Finally{
        $Error.Clear()                                                                                  # Clear error log
    }
} 

Get-Now                                                                                             # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] Processing of PST imports completed - use ExchangeOnline powershell commands to monitor batch progress"

Get-Now
Write-Host  "========================================================" 
Write-Host  "======== $Script:Now Processing Finished =========" 
Write-Host  "========================================================"

Stop-Transcript                                                                                     # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------