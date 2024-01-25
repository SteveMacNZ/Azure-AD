#requires -version 4
<#
.SYNOPSIS
  Connects to SharePoint Admin URL and returns OneDrive URLs with information
.DESCRIPTION
  Connects to SharePoint Admin URL and returns OneDrive URLs with information on usage and Quotas
.PARAMETER None
  None - Read host input for SharePoint admin URL
.INPUTS None
  None - Read host input for SharePoint admin URL
.OUTPUTS
  Log file and csv files stored in script root directory
.NOTES
  Version:        1.0
  Author:         Steve McIntyre
  Creation Date:  23/5/19
  Purpose/Change: Initial Script
.EXAMPLE
  .\GetOneDrives(SizeandQuota).ps1
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

# Script scopped variables
$Script:Date = Get-Date -Format FileDate                                                    # Date format in yyyymmdd
$Script:LogFile = $PSScriptRoot + "\" + $Script:Date + "_" + $CustName + "_ODSites.log"     # logfile location and name
$Script:CSVFile = $PSScriptRoot + "\" + $Script:Date + "_" + $CustName + "_ODSites.csv"     # CSV Export location and name

#-----------------------------------------------------------[Functions]------------------------------------------------------------
# Start Transcriptions
Function Start-Logging{

  Stop-Transcript | out-null                                                                # Ensure no transcription jobs are running
  $ErrorActionPreference = "Continue"                                                       # Set Error Action Handling
  Start-Transcript -path $Script:LogFile -append                                            # Start Transcription append if log exists
  Write-Output ''                                                                           # write Line spacer into Transcription file
  Write-Output ''                                                                           # write Line spacer into Transcription file
  Write-Output '================================================================================================='
  Write-Output ' _____                             _       _   _               _____ _             _           _ '
  Write-Output '|_   _|                           (_)     | | (_)             /  ___| |           | |         | |'
  Write-Output '  | |_ __ __ _ _ __  ___  ___ _ __ _ _ __ | |_ _  ___  _ __   \ `--.| |_ __ _ _ __| |_ ___  __| |'
  Write-Output '  | | |__/ _` | |_ \/ __|/ __| |__| | |_ \| __| |/ _ \| |_ \   `--. \ __/ _` | |__| __/ _ \/ _` |'
  Write-Output '  | | | | (_| | | | \__ \ (__| |  | | |_) | |_| | (_) | | | | /\__/ / || (_| | |  | ||  __/ (_| |'
  Write-Output '  \_/_|  \__,_|_| |_|___/\___|_|  |_| .__/ \__|_|\___/|_| |_| \____/ \__\__,_|_|   \__\___|\__,_|'
  Write-Output '                                    | |                                                          '
  Write-Output '                                    |_|                                                          '
  Write-Output '================================================================================================='   
  
  Write-Output ''                                                                           # write Line spacer into Transcription file
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Script Execution goes here

Start-Logging                                                                               # Call Start-Logging function to start transcription

# Prompt for Tenant Name
$TenantName = Read-Host "Enter the name of the Tenant (e.g. cinz)"                          # prompt for tenant name
$TenantUrl = "https://$TenantName-admin.sharepoint.com"                                     # Create SharePoint Admin URL based on $TenantName
Write-Output "Connecting to $TenantURL now"                                                 # Write Status Update to Transcription file

# Connect to SharePoint 
Connect-SPOService -Url $TenantUrl                                                          # Connect to SharePoint PowerShell Session

# Get all Personal Site collections and export to a Text file
$OneDriveSites = Get-SPOSite -Template "SPSPERS" -Limit ALL -includepersonalsite $True

# Loop through Personal OneDrives and return Usage and Quotas
$Result=@()                                                                                 # Get storage quota of each site
Foreach($Site in $OneDriveSites)
{
    $Result += New-Object PSObject -property @{
    URL = $Site.URL
    Owner= $Site.Owner
    Usage_inMB = $Site.StorageUsageCurrent
    StorageQuota_inGB = $Site.StorageQuota/1024
    StorageQuotaWarningLevel_inGB = $Site.StorageQuotaWarningLevel/1024
    ResourceQuota = $Site.ResourceQuota
    ResourceQuotaWarningLevel = $Site.ResourceQuotaWarningLevel
    }
}
 
$Result | Format-Table                                                                      # Format results as a table

$Result | Select-Object Owner,Url,Usage_inMB,StorageQuota_inGB,StorageQuotaWarningLevel_inGB,ResourceQuota,ResourceQuotaWarningLevel |Export-Csv $Script:CSVFile -NoTypeInformation   # Export the data to CSV

Write-Output "Export of OneDrives completed File saved as $Script:CSVFile."

# Stop Logging
Stop-Transcript