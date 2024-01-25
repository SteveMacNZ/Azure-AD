<#
.SYNOPSIS
  Register-PIMAADGroupsv2
.DESCRIPTION
  what does this script do? extended description 
.PARAMETER None
  None
.INPUTS
  What Inputs  
.OUTPUTS
  What outputs
.NOTES
  Version:        2.0.0.0
  Author:         Steve McIntyre
  Creation Date:  DD/MM/20YY
  Purpose/Change: Re-written to use Graph API to create the PIM Role group and assign eligable role permission
.LINK
  None
.EXAMPLE
  ^ . Register-PIMAADGroupsv2.ps1
  does what with example of cmdlet
  Register-PIMAADGroupsv2.ps1

#>

#requires -version 4
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
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Identity.Governance

#& Includes - Scripts & Modules
. .\Get-CommonFunctions.ps1                                 # Include Common Functions

#endregion
#region -------------------------------------------------------[Declarations]------------------------------------------------------

# Script sourced variables for General settings and Registry Operations
$Script:Date        = Get-Date -Format yyyy-MM-dd                                   # Date format in yyyy-mm-dd
$Script:Now         = ''                                                            # script sourced veriable for Get-Now function
$Script:ScriptName  = 'Register-PIMAADGroupsv2'                                     # Script Name used in the Open Dialogue
$Script:dest        = "$PSScriptRoot\Exports"                                       # Destination path
$Script:LogDir      = "$PSScriptRoot\Logs"                                          # Logdir for Clear-TransLogs function for $PSScript Root
$Script:LogFile     = $Script:LogDir + "\" + $Script:Date + "_" + $env:USERNAME + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:BatchName   = ''                                                            # Batch name variable placeholder
$Script:GUID        = '37c71877-303b-4b5f-b57d-76f2827f5e39'                        # Script GUID
  #^ Use New-Guid cmdlet to generate new script GUID for each version change of the script
[version]$Script:Version  = '2.0.0.0'                                               # Script Version Number
$Script:Client      = ''                                                            # Set Client Name - Used in Registry Operations
$Script:WHO         = whoami                                                        # Collect WhoAmI
$Script:Desc        = ""                                                            # Description displayed in Get-ScriptInfo function
$Script:Desc2       = ""                                                            # Description2 displayed in Get-ScriptInfo function
$Script:PSArchitecture = ''                                                         # Place holder for x86 / x64 bit detection

#^ File Picker / Folder Picker Setup
$Script:File  = ''                                                                  # File var for Get-FilePicker Function
$Script:FPDir       = '$PSScriptRoot'                                               # File Picker Initial Directory
$Script:FileTypes   = "Text files (*.txt)|*.txt|CSV File (*.csv)|*.csv|All files (*.*)|*.*" # File types to be listed in file picker
$Script:FileIndex   = "2"                                                           # What file type to set as default in file picker (based on above order)

$Script:MgDirRoles  = [System.Collections.ArrayList]@()                             # Array list for all Graph API Directory Roles
$Script:MgUsers     = [System.Collections.ArrayList]@()                             # Array list for all Users for Owner Assignment
$Script:PIMObjs     = [System.Collections.ArrayList]@()                             # Array list for PIM Role Assignments Class objects

#endregion
#region --------------------------------------------------------[Hash Tables]------------------------------------------------------

#& any script specific hash tables that are not included in Get-CommonFunctions.ps1

#endregion
#region -------------------------------------------------------[Functions]---------------------------------------------------------

#& any script specific funcitons that are not included in Get-CommonFunctions.ps1

#endregion
#region ------------------------------------------------------------[Classes]-------------------------------------------------------------

#& any script specific classes that are not included in Get-CommonFunctions.ps1

# PIM Group Assignment Class
Class PIMObj{
  # $pimresult = [PIMObj]::new($GroupName,$OwnerName,$OwnerId,$GroupGUID,$PIMRole,$PIMRoleID,$Schedule,$Status)           # creates a new class object
  # $Script:PIMObjs.add($pimresult) | Out-Null                                 # writes the class object to the Class array
  # $Script:PIMObjs | Export-Csv -Path $PIMReport -NoTypeInformation           # writes the class array out to CSV file
  [String]$GroupName
  [String]$OwnerName
  [String]$OwnerId
  [String]$GroupGUID
  [String]$PIMRole
  [String]$PIMRoleID
  [String]$Schedule
  [String]$Status
    
  # constructor
  PIMObj([String]$GroupName,[String]$OwnerName,[String]$OwnerId,[String]$GroupGUID,[String]$PIMRole,[String]$PIMRoleID,[String]$Schedule,[String]$Status){
    $this.GroupName   = $GroupName
    $this.OwnerName   = $OwnerName
    $this.OwnerId     = $OwnerId
    $this.GroupGUID   = $GroupGUID
    $this.PIMRole     = $PIMRole
    $this.PIMRoleID   = $PIMRoleID
    $this.Schedule    = $Schedule
    $this.Status      = $Status
  } 
}


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

Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory,Group.ReadWrite.All"        # Connect to MS Grpah via PowerShell using Modern Auth

Write-Host ""
Write-InfoMsg "Enumerating Directory Roles"
$Script:MgDirRoles = Get-MgDirectoryRoleTemplate | Select-Object DisplayName, Id | Sort-Object DisplayName
$PIMRoles = $Script:dest + "\" + $Script:Date + "_PIM_All_Directory_Roles.csv"          # CSV output of PIM Assignments
$Script:MgDirRoles | Export-Csv -Path $PIMRoles -NoTypeInformation                      # Exports list of all Directory Management Role Templates
Write-Host ""

Write-InfoMsg "Enumerating Users for group Owner lookups. This may take a few minutes. please wait...."
$Script:MgUsers = Get-MGUser | Select-Object DisplayName, Id, UserPrincipalName | Sort-Object DisplayName
Write-Host ""

Get-FilePicker                                                                          # Prompt user for input file

$PIMRoles = Import-csv $Script:File -Delimiter ","                                      # Import CSV file for processing
$counter = 0                                                                            # Init counter
$maximum = $PIMRoles.Count                                                              # number of items to be processed

Write-InfoHighlightedMsg "$maximum Scurity Group / PIM Role assignment Objects found"
Write-Host ""
Foreach ($PRole in $PIMRoles)  {
  Write-Host ""
  # Display progress bar if more than 1 record
  If ($maximum -gt 1){
    $counter++
    $percentCompleted = $counter * 100 / $maximum
    $message = '{0:p1} completed, processing {1}.' -f ( $percentCompleted/100), $PRole.DisplayName
    Write-Progress -Activity 'I am busy' -Status $message -PercentComplete $percentCompleted
  }
  
  Write-InfoMsg "processing PIM Security Group & Role assignment for $($PRole.DisplayName)"

  $Owner          = $PRole.Owner                                                        # Object ID of user object who owns the group
  $Description    = $PRole.Description                                                  # Description for Security Group
  $DisplayName    = $PRole.DisplayName                                                  # Display Name of Security Group
  $MailNickname   = $PRole.mailNickname                                                 # mail nickname of security group
  $PIMRole        = $PRole.RoleName                                                     # Name of the PIM Role wanting to assign
  
  # Create the PIM Role Security group in Entra ID via Graph API
  Try{
    Write-Host "Finding $Owner in the user array list..." -ForegroundColor White
    
    $OwnerExists = $Script:MgUsers | Where-Object {$_.DisplayName -eq "$Owner"}
    If ($OwnerExists){
      Write-InfoMsg "Owner Found: $($OwnerExists.DisplayName) with ID $($OwnerExists.Id)"
      $OwnerId    = "$($OwnerExists.Id)"
      $gpparams = @{
        description = "$Description"
        displayName = "$DisplayName"
        mailEnabled = $false
        mailNickname = "$MailNickname"
        securityEnabled = $true
        isAssignableToRole = $true
        "owners@odata.bind" = @(
          "https://graph.microsoft.com/v1.0/users/$OwnerId"
        )
      }

      Write-InfoMsg "Creating PIM Security Group"
      $PIMGP = New-MgGroup -BodyParameter $gpparams
      #$PIMGP | Format-List                                                              #! [Debug] line to display info of the created group
      Write-Host ""
    }
    else{
      Write-WarningMsg "Unable to locate user to assign as group owner - creating group without owner"
      $OwnerId    = "00000000-0000-0000-0000-000000000000"
      $gpparams = @{
        description = "$Description"
        displayName = "$DisplayName"
        mailEnabled = $false
        mailNickname = "$MailNickname"
        securityEnabled = $true
        isAssignableToRole = $true
      }

      Write-InfoMsg "Creating PIM Security Group"
      $PIMGP = New-MgGroup -BodyParameter $gpparams
      #$PIMGP | Format-List                                                              #! [Debug] line to display info of the created group
      Write-Host ""
    } 
  }
  Catch{
    Write-ErrorMsg "Unable to create security group: $DisplayName"
    Write-Host $PSItem.Exception.Message -ForegroundColor RED                           # Error message details
  }
  Finally{
    $Error.Clear()                                                                      # Clear error log
  }
  Write-Host ""

  Write-Host "Waiting for 15 seconds for group creation to complete in Entra ID" -ForegroundColor Yellow
  Start-Sleep -Seconds 15                                                               #^ Required as without this the PIM role assigment may fail as the group hasn't fully completed
  Write-Host "" 

  Write-Host "Finding $PIMRole in the role array list..." -ForegroundColor White
  $RoleId = $Script:MgDirRoles | Where-Object {$_.DisplayName -eq "$PIMRole"}
  $filter = 'DisplayName eq ' + '"' + $DisplayName + '"'
  $GPExists = Get-MgGroup -Filter "$filter"
  If (!($RoleId -or $GPExists)){
    Write-ErrorMsg "Unknown Role or Group Error - Skipping creation of the PIM Role Assignment....."
    $GroupGUID    = "00000000-0000-0000-0000-000000000000"                              # Set Group ID to blank for reporting
    $PIMRoleDefId = "00000000-0000-0000-0000-000000000000"                              # Set Group ID to blank for reporting
    $PIMSchedule  = "Unknown"
    $PIMRoleState = "Unknown"
  }
  else{
    # Create the PIM group assignment
    Try{
      Write-InfoMsg "Assigning PIM role eligability to $DisplayName"
      $GroupGUID = $PIMGP.Id
      $PIMRoleDefId = $RoleId.Id
      $params = @{
        action = "AdminAssign"
        justification = "Assign $DisplayName eligibility to $PIMRole"
        roleDefinitionId = $PIMRoleDefId
        directoryScopeId = "/"
        principalId = $GroupGUID
        scheduleInfo = @{
          startDateTime = Get-Date
          expiration = @{
          endDateTime = (Get-Date).AddYears(1)
          type = "AfterDateTime"
          }
        }
      }    
      
      $roleAssignment = New-MgRoleManagementDirectoryRoleEligibilityScheduleRequest -BodyParameter $params
      $roleAssignment | Format-List                                           #! [Debug] line to display info of the created group
      $PIMSchedule = "$($roleAssignment.Schedule)"                            # Populate Schedule for export to CSV
      $PIMRoleState = "$($roleAssignment.Status)"                             # Populate Status for export to CSV
      Write-Host ""
    }
    Catch{
        Write-ErrorMsg "Summary of the error message"
        Write-Host $PSItem.Exception.Message -ForegroundColor RED             # Error message details
    }
    Finally{
        $Error.Clear()                                                        # Clear error log
    }  
  }
  
  $pimresult = [PIMObj]::new("$DisplayName","$Owner","$OwnerId","$GroupGUID","$PIMRole","$PIMRoleDefId","$PIMSchedule","$PIMRoleState")           # creates a new class object
  $Script:PIMObjs.add($pimresult) | Out-Null                                  # writes the class object to the Class array
  
  # Clear used variables before next loop
  ($Owner,$Description,$DisplayName,$MailNickname,$PIMRole,$OwnerId,$OwnerExists,$PIMGP,$RoleId,$roleAssignment) = $null
  ($pimresult,$filter,$GPExists,$GroupGUID,$PIMRoleDefId,$PIMSchedule,$PIMRoleState) = $null

  Write-Host ""
  Write-Host "---------------------- $counter of $maximum processed ----------------------"
  Write-Host ""
}

$PIMReport = $Script:dest + "\" + $Script:Date + "_PIM_Role_Assignment_Report.csv" # CSV output of PIM Assignments

Write-InfoMsg "Writing class objects to csv file"                           
$Script:PIMObjs | Export-Csv -Path $PIMReport -NoTypeInformation            # writes the class array out to CSV file

# Input / Output comparsion
Write-Host ""
Write-Host '--------------------------------------------------------------------------------'
Write-Host '|                      Input / Output CSV Count Comparsion                     |'
Write-Host '--------------------------------------------------------------------------------'
$OutputObject = Import-csv "$PIMReport" -Delimiter ","        # Read Output for input/output comparsion
$OutputCount = $OutputObject.count
If ($maximum -eq $OutputCount){
  Get-Now
  Write-Host "$Script:Now [COUNTS] CSV Input [$maximum] and Output [$OutputCount] counts match" @chighlight
}
else{
  Get-Now
  Write-Host "$Script:Now [COUNTS] CSV Input [$maximum] and Output [$OutputCount] counts don't match" @cerror
}
Write-Host '--------------------------------------------------------------------------------'
Write-Host ''

Write-Host ''
Get-Now
Write-Host "$Script:Now [INFORMATION] Processing finished with following outputs"
Write-Host "+ $PIMRoles" -ForegroundColor Yellow
Write-Host "+ $PIMReport" -ForegroundColor Yellow 
Write-Host ''                      

Get-Now
Write-Host "================================================================================"  
Write-Host "================= $Script:Now Processing Finished ====================" 
Write-Host "================================================================================" 

Stop-Transcript
#endregion
#---------------------------------------------------------[Execution Completed]----------------------------------------------------------