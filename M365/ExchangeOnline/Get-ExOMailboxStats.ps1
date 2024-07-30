#region ------------------------------------------------------[Synopsis]--------------------------------------------------
<#
.SYNOPSIS
  Prompts for a CSV input file and connects to Exchange Online and collects mailbox statitics and exports results to CSV
.DESCRIPTION
  Prompts for a CSV input file and connects to Exchange Online and collects mailbox statitics and exports results to CSV 
.PARAMETER None
  None
.INPUTS
  CSV Input file with the following fields:
  EmailAddress      Email Address or UserPrincipal Name (Email preferred)
  DisplayName       Display Name of the User
.OUTPUTS
  CSV output
.NOTES
  Version:        1.0.0.0
  Author:         Steve McIntyre
  Creation Date:  30/07/2024
  Purpose/Change: Initial Release
.LINK
  None
.EXAMPLE
  ^ . Get-ExOMailboxStats.ps1
  Prompts for a CSV input file and connects to Exchange Online and collects mailbox statitics and exports results to CSV
  Get-ExOMailboxStats.ps1

#>
#endregion

#requires -version 7 -Modules ExchangeOnlineManagement
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
Import-Module ExchangeOnlineManagement

#& Includes - Scripts & Modules
. .\Get-CommonFunctions.ps1                                                         # Include Common Functions

#endregion
#region -------------------------------------------------------[Declarations]------------------------------------------------------

# Script sourced variables for General settings and Registry Operations
$Script:Date        = Get-Date -Format yyyy-MM-dd                                   # Date format in yyyy-mm-dd
$Script:Now         = ''                                                            # script sourced veriable for Get-Now function
$Script:ScriptName  = 'Get-ExOMailboxStats'                                         # Script Name used in the Open Dialogue
$Script:dest        = "$PSScriptRoot\Exports"                                       # Destination path
$Script:LogDir      = "$PSScriptRoot\Logs"                                          # Logdir for Clear-TransLogs function for $PSScript Root
$Script:LogFile     = $Script:LogDir + "\" + $Script:Date + "_" + $env:USERNAME + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:BatchName   = ''                                                            # Batch name variable placeholder
$Script:GUID        = '17325a23-2336-4518-a9b4-8d5d311296b0'                        # Script GUID
  #^ Use New-Guid cmdlet to generate new script GUID for each version change of the script
[version]$Script:Version  = '1.0.0.0'                                               # Script Version Number
$Script:Client      = ''                                                            # Set Client Name - Used in Registry Operations
$Script:WHO         = whoami                                                        # Collect WhoAmI
$Script:Desc        = "Exchange Online (ExO) Mailbox Statistics Report"             # Description displayed in Get-ScriptInfo function
$Script:Desc2       = "Gets Mailbox Statistics for a group of users from CSV Input" # Description2 displayed in Get-ScriptInfo function
$Script:PSArchitecture = ''                                                         # Place holder for x86 / x64 bit detection

#^ File Picker / Folder Picker Setup
$Script:File  = ''                                                                  # File var for Get-FilePicker Function
$Script:FPDir       = '$PSScriptRoot'                                               # File Picker Initial Directory
$Script:FileTypes   = "Text files (*.txt)|*.txt|CSV File (*.csv)|*.csv|All files (*.*)|*.*" # File types to be listed in file picker
$Script:FileIndex   = "2"                                                           # What file type to set as default in file picker (based on above order)

$Script:CSVOutput   = $Script:dest + "\" + $Script:Date + "_ExchangeOnline_Mbx_Stats.csv" # CSV Output file locatation and name

#endregion
#region --------------------------------------------------------[Hash Tables]------------------------------------------------------

#& any script specific hash tables that are not included in Get-CommonFunctions.ps1

#endregion
#region -------------------------------------------------------[Functions]---------------------------------------------------------

#& any script specific funcitons that are not included in Get-CommonFunctions.ps1
Function Test-ExchangeOnline{
  # Check for current open O365 sessions and allow the admin to either use the existing session or create a new one
  $EXOsession = Get-ConnectionInformation
  if($null -ne $EXOsession){
    $a = Read-Host "An open session to Exchange Online PowerShell already exists. Do you want to use this session?  Enter y to use the open session, anything else to close and open a fresh session."
    if($a.ToLower() -eq 'y'){
      Write-SuccessMsg "Using existing Exchange Online Powershell session." 
      return
    }
    else{
      Write-Host "Disconnecting from open Exchange Online session" -ForegroundColor Yellow
      Disconnect-ExchangeOnline -Confirm:$false
      Write-Host "Connecting to Exchange Online"
      Connect-ExchangeOnline -ShowBanner:$false
    }  
  }
  else{
    Write-Host "Connecting to Exchange Online"
    Connect-ExchangeOnline -ShowBanner:$false
  }    
}

#endregion
#region ------------------------------------------------------------[Classes]-------------------------------------------------------------

#& any script specific classes that are not included in Get-CommonFunctions.ps1

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

Write-InfoMsg "Please select CSV file from file explorer popup window and click open"
Get-FilePicker                                                                          # Prompt for CSV Input file
Test-ExchangeOnline                                                                     # Test for active Exo Session / Connect to Exchange Online

$Mailboxes = Import-csv $Script:File -Delimiter ","
$counter = 0
$maximum = $Mailboxes.Count                                                             # number of items to be processed

Write-InfoHighlightedMsg "$maximum mailbox Objects found"
Write-Host ""
Foreach ($Mbx in $Mailboxes)  {
  Write-Host ""
  # Display progress bar if more than 1 record
  If ($maximum -gt 1){
    $counter++
    $percentCompleted = $counter * 100 / $maximum
    $message = '{0:p1} completed, processing {1}.' -f ( $percentCompleted/100), $Mbx.DisplayName
    Write-Progress -Activity 'I am busy' -Status $message -PercentComplete $percentCompleted
  }
  
  Write-InfoMsg "processing Exchange Online Mailbox Statistics for $($Mbx.DisplayName)"

  Try{
    Write-InfoMsg "Getting $($Mbx.DisplayName) Mailbox Stats"
    $MbxStats = Get-MailboxStatistics -Identity $Mbx.EmailAddress | Select-Object DeletedItemCount, ItemCount ,TotalDeletedItemSize, TotalItemSize, LastInteractionTime, MailboxTypeDetail
    # try stuff
  }
  Catch{
    Write-ErrorMsg "Unable to return mailbox statistics for $($Mbx.DisplayName)"
    Write-Host $PSItem.Exception.Message -ForegroundColor RED                           # Error message details
  }
  Finally{
    $Error.Clear()                                                                      # Clear error log
  }

  $Result=@{
    'EmailAddress' = $Mbx.EmailAddress;'DisplayName' = $Mbx.DisplayName; 'MailboxType' = $MbxStats.MailboxTypeDetail; 'LastInteractionTime' = $MbxStats.LastInteractionTime;
    'DeletedItemCount' = $MbxStats.DeletedItemCount; 'ItemCount' = $MbxStats.ItemCount; 'TotalDeletedItemSize' = $MbxStats.TotalDeletedItemSize;'TotalItemSize' = $MbxStats.TotalItemSize 
  }
  $Results= New-Object PSObject -Property $Result
  $Results | Select-Object EmailAddress,DisplayName,MailboxType,LastInteractionTime,DeletedItemCount,ItemCount,TotalDeletedItemSize,TotalItemSize `
   | Export-Csv -Path $Script:CSVOutput -NoTypeInformation -Append

  ($MbxStats, $Result, $Results) = $null

  Write-Host ""
  Write-Host "---------------------- $counter of $maximum processed ----------------------"
  Write-Host ""
}

Disconnect-ExchangeOnline -Confirm:$false                                               # Disconnect from ExO with no prompts

# Input / Output comparsion
Write-Host ""
Write-Host '--------------------------------------------------------------------------------'
Write-Host '|                      Input / Output CSV Count Comparsion                     |'
Write-Host '--------------------------------------------------------------------------------'
$OutputObject = Import-csv $Script:CSVOutput -Delimiter ","                             # Read Output for input/output comparsion
$OutputCount = $OutputObject.count
If ($maximum -eq $OutputCount){
  Get-Now
  Write-InfoHighlightedMsg "[COUNTS] CSV Input [$maximum] and Output [$OutputCount] counts match"
}  
else{
  Get-Now
  Write-ErrorMsg "[COUNTS] CSV Input [$maximum] and Output [$OutputCount] counts don't match"
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
#region ------------------------------------------------------[ExtendedHelp]--------------------------------------------------
<#
^ Enter any extended help items here: (e.g., detailed help on functions, commented code blocks so they sit outside of the main script logic)

#>
#endregion