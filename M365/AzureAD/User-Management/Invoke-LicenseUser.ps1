<#
.SYNOPSIS
  Adds cloud user to cloud licensing group based on CSV input
.DESCRIPTION
  Adds cloud user to cloud licensing group based on CSV input
.PARAMETER <Parameter_Name>
  <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  CSV file with following fields 
  UserPrincipalName Email addresss / UPN of the User
  LIC   License type to be applied to the user e.g. E1,E3,E5,F1
.OUTPUTS
  Log file stored in same location as the CSV file
  CSV output of 
.NOTES
  Version:        1.0
  Author:         Steve McIntyre
  Creation Date:  01/10/19
  Purpose/Change: initial script
.EXAMPLE
  .\Invoke-LicenseUser.ps1
#>

#requires -version 4
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
#$Script:SiteCode = Read-Host "Customer Site Code (e.g. ABC)"                                            # Customer Site Code

# Script scopped variables
$Script:Date                = Get-Date -Format yyyy-MM-dd                                               # Date format in yyyymmdd
$Script:File                = ''                                                                        # File var for Get-FilePicker Function
$Script:ScriptName          = 'Invoke-AddSSPR'                                                          # Script Name used in the Open Dialogue
$Script:LogFile             = $PSScriptRoot + "\" + $Script:Date + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:SSPRGP              = "Invoke-LicenseUser"                                                      # Script scoped variable to the SSRP group
<#
$Script:E1                  = "License_" + $Script:SiteCode + "_Office365_E1"                           # Script scoped variable to the E1 license group
$Script:E3                  = "License_" + $Script:SiteCode + "_Office365_E3"                           # Script scoped variable to the E3 license group
$Script:E5                  = "License_" + $Script:SiteCode + "_Office365_E5"                           # Script scoped variable to the E5 license group
$Script:F1                  = "License_" + $Script:SiteCode + "_Office365_F1"                           # Script scoped variable to the F1 license group
#>
$Script:E1                  = "LIC-M365_E1"                                                             # Script scoped variable to the E1 license group
$Script:E3                  = "LIC-M365_E3"                                                             # Script scoped variable to the E3 license group
$Script:E5                  = "LIC-M365_E5"                                                             # Script scoped variable to the E5 license group
$Script:F1                  = "LIC-M365_F1"                                                             # Script scoped variable to the F1 license group
$Script:E1_OID              = ''                                                                        # Script scoped variable E1 Group ObjectID
$Script:E3_OID              = ''                                                                        # Script scoped variable E3 Group ObjectID
$Script:E5_OID              = ''                                                                        # Script scoped variable E5 Group ObjectID
$Script:F1_OID              = ''                                                                        # Script scoped variable F1 Group ObjectID
$Script:GUID                = 'c8bdc7b4-70a7-4234-a2ad-d84826923634'                                    # Script GUID
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
  Write-Host "========================================================" 
  Write-Host "====== $Script:Now Processing Started ========" 
  Write-Host "========================================================" 
  Write-Host ''
  
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

Function Invoke-GetGroups{
    
    $Script:E1_OID = Get-AzureADGroup -Filter "DisplayName eq '$Script:E1'" | Select-Object -ExpandProperty ObjectID
    $Script:E3_OID = Get-AzureADGroup -Filter "DisplayName eq '$Script:E3'" | Select-Object -ExpandProperty ObjectID
    $Script:E5_OID = Get-AzureADGroup -Filter "DisplayName eq '$Script:E5'" | Select-Object -ExpandProperty ObjectID
    $Script:F1_OID = Get-AzureADGroup -Filter "DisplayName eq '$Script:F1'" | Select-Object -ExpandProperty ObjectID

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

Invoke-GetGroups                                                                                    # Call function to find Azure Licensing group IDs

$LICUsers = Import-csv $Script:File -Delimiter ","                                                  # Load CSV into array
$tc = $LICUsers.count                                                                               # Count number of imports to be completed
$lc = 0

ForEach ($LICUser in $LICUsers) {
    # set up progress notification
    $lc++
    Write-Progress "Processing $tc Objects" -Status "Completed: $lc of $tc. Remaining: $($tc-$lc)" -PercentComplete ($lc/$tc*100)  
    Try {
        Get-Now
        Write-Host "$Script:Now [INFORMATION] Processing "$_.UserPrincipalName"" -ForegroundColor Yellow
        $UsrOID = Get-AzureADUser -ObjectId $_."UserPrincipalName" | Select-Object -ExpandProperty ObjectID # Obtain Users Azure AD ID
        if ($_.LIC -eq 'E1'){
            Add-AzureADGroupMember -ObjectId $Script:E1_OID -RefObjectId $UsrOID                    # Add user to E1 group if they match
        }
        elseif ($_.LIC -eq 'E3'){
            Add-AzureADGroupMember -ObjectId $Script:E3_OID -RefObjectId $UsrOID                    # Add user to E3 group if they match
        }
        elseif ($_.LIC -eq 'E5'){
            Add-AzureADGroupMember -ObjectId $Script:E5_OID -RefObjectId $UsrOID                    # Add user to E5 group if they match       
        }
        elseif ($_.LIC -eq 'F1'){
            Add-AzureADGroupMember -ObjectId $Script:F1_OID -RefObjectId $UsrOID                    # Add user to F1 group if they match
        }
        else {
            Write-Output "An Error Occured"
        } 
    }
    Catch {
        Get-Now
        Write-Host "$Script:Now [ERROR] There was an error adding user "$_.UserPrincipalName" to licensing group" -ForegroundColor Red
        Write-Host $PSItem.Exception.Message -ForegroundColor RED
    }
    Finally{
        $Error.Clear()                                                                              # Clear error log
    }
}    

Get-Now                                                                                             # Get Current Date Time
Write-Host "$Script:Now [INFORMATION] Processing of User licensing completed - users should now be licensed"

Get-Now
Write-Host "========================================================" 
Write-Host "======== $Script:Now Processing Finished =========" 
Write-Host "========================================================"

Stop-Transcript                                                                                     # Stop transcription

#---------------------------------------------------------[Execution Completed]----------------------------------------------------------


