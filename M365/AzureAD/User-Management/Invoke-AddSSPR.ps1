<#
.SYNOPSIS
  Adds cloud user to cloud SSPR group based on CSV input
.DESCRIPTION
  Adds cloud user to cloud SSRP group based on CSV input
.PARAMETER None
  None
.INPUTS
  CSV file with following fields 
  UserPrincipalName Email addresss / UPN of the User
.OUTPUTS
  Transcription log stored in the script root directory
.NOTES
  Version:        1.1
  Author:         Steve McIntyre
  Creation Date:  18/08/2021
  Purpose/Change: Updates for function changes and logic for inital upload to GitHub + additional error handling
  Version:        1.0
  Author:         Steve McIntyre
  Creation Date:  01/10/19
  Purpose/Change: initial script
.LINK
  https://github.com/SteveMacNZ/Azure-M365/tree/main/M365/Azure-AD/User-Management
.EXAMPLE
  .\Invoke-AddSSRP.ps1
#>

#requires -version 4 -Modules AzureAD
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
$Script:ScriptName          = 'Invoke-AddSSPR'                                                          # Script Name used in the Open Dialogue
$Script:LogFile             = $PSScriptRoot + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:SSPRGP              = "Self Service Password Reset"                                             # Script scoped variable to the SSRP group
$Script:SSPRGP_OID          = ''                                                                        # Script scoped variable SSRP Group ObjectID
$Script:GUID                = '27ef61a8-efec-4e48-b3c0-f65fd333194e'                                    # Script GUID
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

#& FilePicker function for selecting input file via explorer window
Function Get-FilePicker {
    Param ()
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.InitialDirectory = $PSScriptRoot                                                         # Sets initial directory to script root
    $ofd.Title            = "Select file for $Script:ScriptName"                                  # Title for the Open Dialogue
    $ofd.Filter           = "Text files (*.txt)|*.txt|CSV File (*.csv)|*.csv|All files (*.*)|*.*" # Display All files / Txt / CSV
    $ofd.FilterIndex      = 2                                                                     # Default to display All files
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
    Finally {
        $Error.Clear()                                                                                  # Clear error log
    }  
    
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

Function Invoke-GetSSPRGroupID {
    Try {
        Get-Now
        Write-Host "$Script:Now [INFORMATION] Looking up SSPR object ID for group: $Script:SSPRGP"
        $Script:SSPRGP_OID = Get-AzureADGroup -Filter "DisplayName eq '$Script:SSPRGP'" | Select-Object -ExpandProperty ObjectID
        Get-Now
        Write-Host "$Script:Now [INFORMATION] Looking up SSPR object ID for group: $Script:SSPRGP" -ForeGroundColor Green
    }
    Catch {
        Get-Now
        Write-Host "$Script:Now [ERROR] There was an error finding group: $Script:SSPRGP" -ForegroundColor Red
        Write-Host $PSItem.Exception.Message -ForegroundColor RED
    }
    Finally {
        $Error.Clear()                                                                                  # Clear error log
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

Invoke-GetSSPRGroupID                                                                               # Call function to find Azure Licensing group IDs

$SSPRUsers = Import-csv $Script:File -Delimiter ","                                                 # Load CSV into array
$tc = $SSPRUsers.count                                                                              # Count number of imports to be completed
$lc = 0

If ($null -ne $Script:SSPRGP_OID){
    Get-Now
    Write-Host "$Script:Now [INFORMATION] SSPR Group exists proceeding with adding users to group"
    ForEach ($SSPRUser in $SSPRUsers) {
        # set up progress notification
        $lc++
        Write-Progress "Processing $tc Objects" -Status "Completed: $lc of $tc. Remaining: $($tc-$lc)" -PercentComplete ($lc/$tc*100)  
        Try {
            Get-Now
            Write-Host "$Script:Now [INFORMATION] Processing "$_.UserPrincipalName"" -ForegroundColor Yellow
            $UsrOID = Get-AzureADUser -ObjectId $_."UserPrincipalName" | Select-Object -ExpandProperty ObjectID # Obtain Users Azure AD ID
            Add-AzureADGroupMember -ObjectId $Script:SSPRGP_OID -RefObjectId $UsrOID                 # Add user to SSRP Group
        }
        Catch {
            Get-Now
            Write-Host "$Script:Now [ERROR] There was an error with the PST import for "$_.UserPrincipalName"" -ForegroundColor Red
            Write-Host $PSItem.Exception.Message -ForegroundColor RED
        }
        Finally{
            $Error.Clear()                                                                           # Clear error log
        }
    }    
}
else {
    Get-Now
    Write-Host "$Script:Now [INFORMATION] SSPR Group does not exist. Existing"
        
}

Get-Now                                                                                             # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] Processing of SSPR Users completed - Users should now be prompted to enrol in SSPR"

Get-Now
Write-Host  "========================================================" 
Write-Host  "======== $Script:Now Processing Finished =========" 
Write-Host  "========================================================"

Stop-Transcript                                                                                     # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------