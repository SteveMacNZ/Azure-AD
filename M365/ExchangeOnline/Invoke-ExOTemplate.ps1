<#
.SYNOPSIS
  Synopsis of the script
.DESCRIPTION
  Description of the script
.PARAMETER None
  None
.INPUTS
  None
.OUTPUTS
  Log file for transcription logging
.NOTES
  Version:        #.#
  Author:         <name>     
  Creation Date:  dd/mm/yy
  Purpose/Change: Initial Script
.LINK
  None
.EXAMPLE
  .\Invoke-ExOTemplate.ps1
  Connects to Exchange Online and <what>
 
#>


#Requires -Modules ExchangeOnlineManagement
#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  #Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins
#Import-Module ActiveDirectory

#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Script scopped variables
$Script:Date                = Get-Date -Format yyyy-MM-dd                                               # Date format in yyyymmdd
$Script:File                = ''                                                                        # File var for Get-FilePicker Function
$Script:ScriptName          = 'Invoke-ExOTemplate'                                                      # Script Name used in the Open Dialogue
$Script:LogFile             = $PSScriptRoot + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:GUID                = ''                                                                        # Script GUID
#^ Use New-Guid cmdlet to generate new script GUID for each version change of the script
$Script:LogDir              = ""                                                                        # Output directory pathfor exports
#$Script:LogDir              = "\\server\share\sub-folder"                                              # Output directory pathfor exports
$Script:BatchFolder         = ''                                                                        # Placeholder for Batchfolder name
$Script:BFolder             = ''                                                                        # Placeholder for LogDir + BatchFolder path

#-----------------------------------------------------------[Hash Tables]-----------------------------------------------------------

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

#& FilePiucker function for selecting input file via explorer window
Function Get-FilePicker {
    Param ()
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.InitialDirectory = $PSScriptRoot                                                         # Sets initial directory to script root
    $ofd.Title            = "Select file for $Script:ScriptName"                                  # Title for the Open Dialogue
    $ofd.Filter           = "Text files (*.txt)|*.txt|CSV File (*.csv)|*.csv|All files (*.*)|*.*" # Display All files / Txt / CSV
    $ofd.FilterIndex      = 1                                                                     # Default to display All files
    $ofd.RestoreDirectory = $true                                                                 # Reset the directory path
    #$ofd.ShowHelp         = $true                                                                 # Legacy UI              
    $ofd.ShowHelp         = $false                                                                # Modern UI
    if($ofd.ShowDialog() -eq "OK") { $ofd.FileName }
    $Script:File = $ofd.Filename
}

#& Test-Exchange function for testing for connection to On-Prem Exchange connecting if no connection is found
Function Test-Exchange{
    try {
      #^ Attempts to return Exchange Servers to verify connection to Exchange OnPrem Tools
      $IsExchangeShell = Get-ExchangeServer -ErrorAction SilentlyContinue
    }
     
    Catch {}
    
    if ($null -eq $IsExchangeShell){
      Get-Now
      Write-Host "$Script:Now [INFORMATION] Exchange - is not connected - Connecting...." -ForegroundColor Magenta
      Import-Module "$env:exchangeinstallpath\Bin\RemoteExchange.ps1"
      Connect-ExchangeServer -auto -ClientApplication:ManagementShell
    }
    else {
      Get-Now
      write-host "$Script:Now [INFORMATION] Already connected to Exchange - Proceeding" -ForegroundColor Green 
    }
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

#& TestPath function for testing and creating directories
Function Invoke-TestPath{
    [CmdletBinding()]
    param (
        #^ Path parameter for testing/creating destination paths
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [String]
        $ParamPath
    )
    Try{
        # Check to see if the report location exists, if not create it
        if ((Test-Path -Path $ParamPath -PathType Container) -eq $false){
            Get-Now
            Write-Host "$Script:Now [INFORMATION] Destination Path $($ParamPath) does not exist: creating...." -ForegroundColor Magenta -BackgroundColor White
            New-Item $ParamPath -ItemType Directory | Out-Null
            Get-Now
            Write-Verbose "$Script:Now [INFORMATION] Destination Path $($ParamPath) created"
        }
    }  
    Catch{
        #! Error handling for folder creation 
        Get-Now
        Write-Host "$Script:Now [Error] Error creating directories"
        Write-Host $PSItem.Exception.Message
        Stop-Transcript
        Break
    }
}


#-----------------------------------------------------------[Execution]------------------------------------------------------------
<#
? ---------------------------------------------------------- [NOTES:] -------------------------------------------------------------
& Best veiwed and edited with Microsoft Visual Studio Code with colorful comments extension
* Requires ExchangeOnlineManagement V2 Module
* If also Connecting to Exchange On-Prem management tools need to be installed on the host running the script
* Transcription logging formatting use Get-Now before write-host to return current timestamp into $Scipt:Now variable
  Write-Host "$Script:Now [INFORMATION] Information Message"
  Write-Host "$Script:Now [WARNING] Warning Message"
  Write-Host "$Script:Now [ERROR] Error Message"
? ---------------------------------------------------------------------------------------------------------------------------------
#>

# Script Execution goes here
Start-Logging                                                                                       # Start Transcription logging
#Test-Exchange                                                                                       # Test connection to Exchange (On Prem)
Test-ExchangeOnline                                                                                 # Test connection to Exchange Online

Write-Host ""
Get-FilePicker                                                                                      # Call FilePicker Function
Get-Now                                                                                             # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] $Script:File has been selected for processing" -ForegroundColor Magenta
Write-Host ""

#* Sets Batch Name to be name of the file selected in FilePicker function and uses in destination folder structures
$BatchNameTemp      = $Script:File.split("\")[-1]
$Script:BatchFolder = $BatchNameTemp.substring(0,($BatchNameTemp.length-4))
$Script:BFolder     = $Script:LogDir + "\" + $Script:BatchFolder


# Test and create folder structure
Invoke-TestPath -ParamPath $Script:LogDir
Invoke-TestPath -ParamPath $Script:BFolder


<#
    !! Do Stuff Here
#>

Get-Now
Write-Host "$Script:Now [INFORMATION] Files Successfully Exported to $Script:BFolder" -ForeGroundColor Green

Get-Now
Write-Output  "========================================================" 
Write-Output  "======== $Script:Now Processing Finished =========" 
Write-Output  "========================================================"

Stop-Transcript                                                                                     # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------