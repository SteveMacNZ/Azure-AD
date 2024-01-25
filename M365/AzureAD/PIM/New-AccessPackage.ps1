<#
.SYNOPSIS
  what does this script do?
.DESCRIPTION
  what does this script do? extended description 
.PARAMETER None
  None
.INPUTS
  What Inputs  
.OUTPUTS
  What outputs
.NOTES
  Version:        1.0.0.x
  Author:         Steve McIntyre
  Creation Date:  DD/MM/20YY
  Purpose/Change: Initial Release
.LINK
  None
.EXAMPLE
  ^ . New-AccessPackage.ps1
  does what with example of cmdlet
  Invoke-What.ps1

#>

#requires -version 4 -Modules Microsoft.Graph.Identity.Governance, Microsoft.Graph.Beta.Identity.Governance
#region ------------------------------------------------------[Script Parameters]--------------------------------------------------

Param (
  #Script parameters go here
)

#endregion
#region ------------------------------------------------------[Initialisations]----------------------------------------------------

#& Global Error Action
#$ErrorActionPreference = 'SilentlyContinue'

#& Module Imports
Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Identity.Governance, Microsoft.Graph.Beta.Identity.Governance

#& Includes - Scripts & Modules
. .\Get-CommonFunctions.ps1                                 # Include Common Functions

#endregion
#region -------------------------------------------------------[Declarations]------------------------------------------------------

# Script sourced variables for General settings and Registry Operations
$Script:Date        = Get-Date -Format yyyy-MM-dd                                   # Date format in yyyy-mm-dd
$Script:Now         = ''                                                            # script sourced veriable for Get-Now function
$Script:ScriptName  = 'New-AccessPackage'                                           # Script Name used in the Open Dialogue
$Script:dest        = "$PSScriptRoot\Exports"                                       # Destination path
$Script:LogDir      = "$PSScriptRoot\Logs"                                          # Logdir for Clear-TransLogs function for $PSScript Root
$Script:LogFile     = $Script:LogDir + "\" + $Script:Date + "_" + $env:USERNAME + "_" + $Script:ScriptName + ".log"    # logfile location and name
$Script:BatchName   = ''                                                            # Batch name variable placeholder
$Script:GUID        = 'fd9dcdd0-e9a7-4b25-927c-dad0244e05df'                        # Script GUID
  #^ Use New-Guid cmdlet to generate new script GUID for each version change of the script
[version]$Script:Version  = '0.0.0.0'                                               # Script Version Number
$Script:Client      = ''                                                            # Set Client Name - Used in Registry Operations
$Script:WHO         = whoami                                                        # Collect WhoAmI
$Script:Desc        = ""                                                            # Description displayed in Get-ScriptInfo function
$Script:Desc2       = ""                                                            # Description2 displayed in Get-ScriptInfo function
$Script:PSArchitecture = ''                                                         # Place holder for x86 / x64 bit detection

#^ File Picker / Folder Picker Setup
[System.IO.FileInfo]$Script:File  = ''                                              # File var for Get-FilePicker Function
$Script:FPDir       = '$PSScriptRoot'                                               # File Picker Initial Directory
$Script:FileTypes   = "Text files (*.txt)|*.txt|CSV File (*.csv)|*.csv|All files (*.*)|*.*" # File types to be listed in file picker
$Script:FileIndex   = "2"                                                           # What file type to set as default in file picker (based on above order)

#endregion
#region --------------------------------------------------------[Hash Tables]------------------------------------------------------

#& any script specific hash tables that are not included in Get-CommonFunctions.ps1

#endregion
#region -------------------------------------------------------[Functions]---------------------------------------------------------

#& any script specific funcitons that are not included in Get-CommonFunctions.ps1


#endregion
#region ------------------------------------------------------------[Classes]-------------------------------------------------------------

#& any script specific classes that are not included in Get-CommonFunctions.ps1

<#
# Example Class - constuct and usage
Class ClassName{
  # $classresult = [ClassName]::new("$WhatString","$WhatINT","$WhatBool")           # creates a new class object
  # $Script:ClassArray.add($classresult) | Out-Null                                 # writes the class object to the Class array
  # $Script:ClassArray | Export-Csv -Path $ClassReport -NoTypeInformation           # writes the class array out to CSV file
  [String]$WhatString
  [INT]$WhatINT 
  [Bool]$WhatBool
  
  # constructor
  ClassName([String]$WhatString, [INT]$WhatINT, [Bool]$WhatBool){
    $this.WhatString = $WhatString
    $this.WhatINT = $WhatINT
    $this.WhatBool = $WhatBool
  } 
}
#>

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

# Connect to Graph API
Connect-MgGraph -Scopes "EntitlementManagement.ReadWrite.All,Group.ReadWrite.All"

#* Prompt to create new catalog item
$NewCat = Show-ConsoleDialog -Message 'Do you want to create a new Catalog Item?' -Title 'Create new Catalog Item' -Choice 'Yes','No'
switch ($NewCat){
  'Yes' { 
    Write-InfoMsg "Creating new Catalog Item"
    $DisName = Read-host "Enter the display name of the catalog"
    $Desc = Read-Host "Enter Catalog Description"
    $result = Show-ConsoleDialog -Message 'Should the Catalog be visible to external users?' -Title 'Catalog Visibility' -Choice 'Yes','No'
    switch ($result){
      'Yes' { $extVis = $true }
      'No'  { $extVis = $false }
    }
    $cat_params = @{
      displayName = "$DisName"
      description = "$Desc"
      state = "published"
      isExternallyVisible = $extVis
    }

    $CATA = New-MgEntitlementManagementCatalog -BodyParameter $cat_params               # Create new Catalog Item

    $CATA | Format-List                                                                 # Display details of created Catalog for Transcription file    

  }
  'No'  { Write-InfoMsg "Skipping Catalog creation" }
}

#* Prompt to add resources to a catalog using CSV input file
$AddCatRes = Show-ConsoleDialog -Message 'Do you want to add resources to a catalog [CSV Required]?' -Title 'Add Catalog Resources' -Choice 'Yes','No'
switch ($AddCatRes){
  'Yes' { 
    Write-InfoMsg "Adding Resources to Catalog"
    Get-FilePicker                                                                      # Prompt user for input file
    
    $ResToAdd = Import-csv $Script:File -Delimiter ","                                  # Import CSV file for processing
    $counter = 0                                                                        # Init counter
    $maximum = $ResToAdd.Count                                                          # number of items to be processed
    Write-InfoHighlightedMsg "$maximum Resource additions found"
    Write-Host ""
    Foreach($Res in $ResToAdd){
      # Display progress bar if more than 1 record
      If ($maximum -gt 1){
        $counter++
        $percentCompleted = $counter * 100 / $maximum
        $message = '{0:p1} completed, processing {1}.' -f ( $percentCompleted/100), $PRole.DisplayName
        Write-Progress -Activity 'I am busy' -Status $message -PercentComplete $percentCompleted
      }
    
      $Catalog      = $Res.Catalog                                                      # Object ID of user object who owns the group
      $Resource     = $Res.ResourceToAdd                                                # Description for Security Group
      Write-InfoMsg "processing addition of $Resource to $Catalog"
      Write-Host "Searching Entra ID for $Resource" -ForegroundColor Yellow
      $gfilter      = "(displayName eq '" + $Resource + "')"                            # Filter for group search
      Write-Host "[DEBUG] Search filter: $($gfilter) will be used"
      Write-Host "[DEBUG] Get-MgGroup -Filter $($gfilter) will be run"
      $g = Get-MgGroup -Filter "$gfilter"                                                 # Return Entra ID Group details
      if ($null -eq $g) {throw "no group" } else {Write-SuccessMsg "$Resource located"}

      Write-Host "Searching Entra ID for $Catalog" -ForegroundColor Yellow
      $cfilter       = "(displayName eq '" + $Catalog + "')"                            # Filter for catalog search
      Write-Host "[DEBUG] Search filter: $($cfilter) will be used"
      Write-Host "[DEBUG] Get-MgBetaEntitlementManagementAccessPackageCatalog -Filter $($cfilter) will be run"
      $CatDets = Get-MgBetaEntitlementManagementAccessPackageCatalog -Filter "$cfilter"   # Return catalog details
       
      if ($null -eq $CatDets) { throw "catalog not found" } else {Write-SuccessMsg "$Catalog located"}
      
      Try{
        Write-InfoMsg "Adding $Resource to Catalog: $Catalog"

        $res_params = @{
          catalogId = "$($CatDets.Id)"
          requestType = "AdminAdd"
          accessPackageResource = @{
            originId = $g.Id
            originSystem = "AadGroup"
          }
        }
        
        New-MgBetaEntitlementManagementAccessPackageResourceRequest -BodyParameter $res_params
        Start-Sleep 5
      }
      Catch{
        Write-ErrorMsg "Unable to add $Resouce to $Catalog"
        Write-Host $PSItem.Exception.Message -ForegroundColor RED                       # Error message details
      }
      Finally{
        $Error.Clear()                                                                  # Clear error log
      }
      
      ($Resource, $Catalog, $g, $gfilter, $CatDets, $cfilter, $res_params ) = $null
    }

  }
'No'  { Write-InfoMsg "Skipping adding resources to catalog" }
}

#* Prompt to add create single or multiple Access packages
$AccPkg = Show-ConsoleDialog -Message 'Do you want to create a single or bulk [CSV Required] access package?' -Title 'Create Access Package' -Choice 'Single','Bulk', 'Cancel'
switch ($AccPkg){
  'Single' { 
    $PkgName = Read-Host "Enter the Name of the Access Package you want to create (e.g., Fujitsu - Professional Services)"
    $PkgDesc = Read-Host "Enter the Description of the Access Package you want to create (e.g, Access Package for Fujitsu Professional Services Staff)"
    $CatName = Read-Host "Enter the Catalog Name that the access package will use (e.g., CAT_Fujitsu)"

    Write-Host "Searching Entra ID for $CatName" -ForegroundColor Yellow
    $cfilter       = "(displayName eq '" + $CatName + "')"                            # Filter for catalog search
    $Catalog = Get-MgBetaEntitlementManagementAccessPackageCatalog -Filter $cfilter   # Return catalog details
    if ($null -eq $Catalog) { throw "catalog not found" } else {Write-SuccessMsg "$CatName located"}
    
    Write-InfoMsg "Creating Access Package: $PkgName"     
    Try{
      Write-InfoMsg ""
      $pkg_params = @{
        displayName = "$PkgName"
        description = "$PkgDesc"
        catalog = @{
            id = $Catalog.id
        }
      }
    $ap = New-MgEntitlementManagementAccessPackage -BodyParameter $pkg_params
    Start-Sleep 5  
    }
    Catch{
      Write-ErrorMsg "Unable to create access package $PkgName"
      Write-Host $PSItem.Exception.Message -ForegroundColor RED                         # Error message details
    }
    Finally{
      $Error.Clear()                                                                    # Clear error log
    }
  }
  'Bulk'  { 
    Get-FilePicker                                                                      # Prompt user for input file
    
    $PkgToAdd = Import-csv $Script:File -Delimiter ","                                  # Import CSV file for processing
    $counter = 0                                                                        # Init counter
    $maximum = $PkgToAdd.Count                                                          # number of items to be processed
    Write-InfoHighlightedMsg "$maximum Access Packages requested"
    Write-Host ""
    Foreach($Pkg in $PkgToAdd){
      # Display progress bar if more than 1 record
      If ($maximum -gt 1){
        $counter++
        $percentCompleted = $counter * 100 / $maximum
        $message = '{0:p1} completed, processing {1}.' -f ( $percentCompleted/100), $PRole.DisplayName
        Write-Progress -Activity 'I am busy' -Status $message -PercentComplete $percentCompleted
      }

      $PkgName      = $Pkg.PackageName                                                  # Name of Access Package
      $PkgDesc      = $Pkg.PackageDescription                                           # Description for Access Package
      $CatName      = $Pkg.Catalog                                                      # Name of the Catalog Item
          
      Write-InfoMsg "processing creation of $PkgName"
      Write-Host "Searching Entra ID for $CatName" -ForegroundColor Yellow
      $cfilter       = "(displayName eq '" + $CatName + "')"                            # Filter for catalog search
      $Catalog = Get-MgBetaEntitlementManagementAccessPackageCatalog -Filter $cfilter   # Return catalog details
      if ($null -eq $Catalog) { throw "catalog not found" } else {Write-SuccessMsg "$CatName located"}

      Write-InfoMsg "Creating Access Package: $PkgName"     
      Try{
        Write-InfoMsg ""
        $pkg_params = @{
          displayName = "$PkgName"
          description = "$PkgDesc"
          catalog = @{
            id = $Catalog.id
          }
        }
        $ap = New-MgEntitlementManagementAccessPackage -BodyParameter $pkg_params
        Start-Sleep 5  
      }
      Catch{
        Write-ErrorMsg "Unable to create access package $PkgName"
        Write-Host $PSItem.Exception.Message -ForegroundColor RED                       # Error message details
      }
      Finally{
        $Error.Clear()                                                                  # Clear error log
      }
      
      ($PkgName, $PkgDesc, $CatName, $Catalog, $cfilter, $pkg_params ) = $null
    }     
}
'Cancel'  { Write-InfoMsg "Skipping creation of Access Package" }
}

<#
# Example foreach loop with input / output count validation
#$What = Import-csv $Script:File -Delimiter ","
$What = "cmdlet to collect required information e.g., Get-ADUser"
$counter = 0
$maximum = $What.Count  # number of items to be processed

Write-InfoHighlightedMsg "$maximum What Objects found"
Write-Host ""
Foreach ($W in $What)  {
  Write-Host ""
  # Display progress bar if more than 1 record
  If ($maximum -gt 1){
    $counter++
    $percentCompleted = $counter * 100 / $maximum
    $message = '{0:p1} completed, processing {1}.' -f ( $percentCompleted/100), $PRole.DisplayName
    Write-Progress -Activity 'I am busy' -Status $message -PercentComplete $percentCompleted
  }

  Write-InfoMsg "processing what for $($W.Value)"

  # doing stuff here
  Try{
    Write-InfoMsg "What is being attempted"
    # try stuff
  }
  Catch{
    Write-ErrorMsg "Summary of the error message"
    Write-Host $PSItem.Exception.Message -ForegroundColor RED             # Error message details
  }
  Finally{
    $Error.Clear()                                                        # Clear error log
  }

  $classresult = [ClassName]::new("$WhatString","$WhatINT","$WhatBool")
  Write-SuccessMsg "$($W.Value) written to class object"
  $Script:ClassResults.add($classresult) | Out-Null

  ("$WhatString","$WhatINT","$WhatBool",$AnyOtherVarsThatNeedtobeCleared) = $null

}

$ClassReport = $Script:dest + "\" + $Script:Date + "_ClassReport.csv"

Write-InfoMsg "Writing class objects to csv file"
$Script:ClassArray | Export-Csv -Path $ClassReport -NoTypeInformation

# Input / Output comparsion
Write-Host ""
Write-Host '--------------------------------------------------------------------------------'
Write-Host '|                      Input / Output CSV Count Comparsion                     |'
Write-Host '--------------------------------------------------------------------------------'
$OutputObject = Import-csv "$ClassReport" -Delimiter ","        # Read Output for input/output comparsion
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