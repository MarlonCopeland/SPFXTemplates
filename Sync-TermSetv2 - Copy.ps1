<#
.SYNOPSIS
    Synchronizes a SharePoint Term Set from a .psd1 source file.

.DESCRIPTION
    This script reads a .psd1 file containing Term Set data (Name, ID, Terms, Labels) and synchronizes it
    to a target SharePoint Term Store. It acts as a "Source of Truth" sync:
    - Adds missing terms.
    - Updates existing terms (Name, Description, Labels).
    - DELETES terms that are not in the source file.
    - Matches terms by ID.

.PARAMETER InputFile
    Path to the source .psd1 file.

.PARAMETER SiteUrl
    The URL of the SharePoint site to connect to.

.PARAMETER TermGroupName
    The name of the Term Group where the Term Set should exist.

.PARAMETER WhatIf
    If specified, runs in simulation mode and outputs what would happen without making changes.

.EXAMPLE
    .\Sync-TermSet.ps1 -InputFile "SSATermSet.psd1" -SiteUrl "https://contoso.sharepoint.com/sites/intranet" -TermGroupName "Organizational Hierarchy" -WhatIf
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,

    [switch]$WhatIf
)

# Helper to log messages
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $Color
}

function Import-TermSetData {
    param([string]$Path)
    $content = Get-Content -Path $Path -Raw
    Invoke-Expression $content
}

function Sync-Terms {
    param (
        [object[]]$SourceTerms,
        [object]$ParentEntity, # Can be TermSet or Term
        [string]$TermStoreId,
        [string]$TermSetId
    )

    # Get existing children of the parent
    # Note: PnP retrieval depends on whether parent is TermSet or Term
    $existingTerms = @()
    if ($ParentEntity.GetType().Name -match "TermSet") {
        # It's a TermSet
        $existingTerms = Get-PnPTerm -TermSet $ParentEntity.Id -TermStore $TermStoreId -Recursive:$false
    }
    else {
        # It's a Term
        $existingTerms = Get-PnPTerm -Term $ParentEntity.Id -TermSet $TermSetId -TermStore $TermStoreId -Recursive:$false
    }

    # Create a lookup for existing terms by ID
    $existingTermsById = @{}
    foreach ($term in $existingTerms) {
        $existingTermsById[$term.Id.ToString()] = $term
    }

    # 1. Process Source Terms (Add/Update)
    foreach ($sourceTerm in $SourceTerms) {
        $sourceId = $sourceTerm.ID
        $sourceName = $sourceTerm.Name
        $sourceDesc = if ($sourceTerm.Description) { $sourceTerm.Description } else { "" }
        $sourceLabels = if ($sourceTerm.Labels) { $sourceTerm.Labels } else { @() }
        
        # Ensure the Name is included in labels for comparison purposes, 
        # though strictly Name is the default label.
        
        if ($existingTermsById.ContainsKey($sourceId)) {
            # UPDATE
            $targetTerm = $existingTermsById[$sourceId]
            Write-Log "Checking Term: $($targetTerm.Name) ($sourceId)" -Color Gray
            
            # Check Name
            if ($targetTerm.Name -ne $sourceName) {
                if ($WhatIf) {
                    Write-Log "  [WhatIf] Rename '$($targetTerm.Name)' to '$sourceName'" -Color Yellow
                }
                else {
                    Write-Log "  Updating Name to '$sourceName'" -Color Yellow
                    Set-PnPTerm -Identity $targetTerm.Id -Name $sourceName -TermStore $TermStoreId
                }
            }

            # Check Description
            if ($targetTerm.Description -ne $sourceDesc) {
                if ($WhatIf) {
                    Write-Log "  [WhatIf] Update Description" -Color Yellow
                }
                else {
                    Write-Log "  Updating Description" -Color Yellow
                    Set-PnPTerm -Identity $targetTerm.Id -Description $sourceDesc -TermStore $TermStoreId
                }
            }

            # Check Labels (Synonyms)
            # We need to fetch labels for the term
            # PnP object might not have labels loaded by default in the list view?
            # Let's assume we need to load them or get them.
            # For efficiency, we might assume they are loaded if we requested them, 
            # but Get-PnPTerm usually returns basic info.
            # Let's load labels explicitly if not in WhatIf (or even in WhatIf to compare).
            
            # To properly compare labels, we need to get them.
            # Using CSOM context to load labels for the current term
            $ctx = Get-PnPContext
            $ctx.Load($targetTerm.Labels)
            $ctx.ExecuteQuery()
            
            $existingLabels = $targetTerm.Labels | ForEach-Object { $_.Value }
            
            # Determine labels to add and remove
            # Source labels usually include the Name. 
            # SharePoint treats the Name as the default label.
            # Other labels are synonyms.
            
            # Labels to add: In Source but not in Existing
            $labelsToAdd = $sourceLabels | Where-Object { $_ -notin $existingLabels }
            
            # Labels to remove: In Existing but not in Source
            # Note: We must NOT remove the label that matches the current Name (Default Label)
            # even if it's somehow missing from source (though source should have it).
            $labelsToRemove = $existingLabels | Where-Object { $_ -notin $sourceLabels -and $_ -ne $sourceName }

            foreach ($label in $labelsToAdd) {
                if ($WhatIf) {
                    Write-Log "  [WhatIf] Add Label: $label" -Color Green
                }
                else {
                    Write-Log "  Adding Label: $label" -Color Green
                    # New-PnPTermLabel doesn't exist? 
                    # Use object method
                    $targetTerm.CreateLabel($label, 1033, $false) # 1033 = English, false = not default
                    $ctx.ExecuteQuery()
                }
            }

            foreach ($label in $labelsToRemove) {
                if ($WhatIf) {
                    Write-Log "  [WhatIf] Remove Label: $label" -Color Red
                }
                else {
                    Write-Log "  Removing Label: $label" -Color Red
                    # Find the label object
                    $labelObj = $targetTerm.Labels | Where-Object { $_.Value -eq $label }
                    if ($labelObj) {
                        $labelObj.DeleteObject()
                        $ctx.ExecuteQuery()
                    }
                }
            }

            # Recurse
            if ($sourceTerm.Terms) {
                Sync-Terms -SourceTerms $sourceTerm.Terms -ParentEntity $targetTerm -TermStoreId $TermStoreId -TermSetId $TermSetId
            }
            elseif ($targetTerm.TermsCount -gt 0) {
                # If source has no children, but target does, we need to check if we should delete target children.
                # We pass empty array to sync to trigger deletion logic.
                Sync-Terms -SourceTerms @() -ParentEntity $targetTerm -TermStoreId $TermStoreId -TermSetId $TermSetId
            }

            # Remove processed ID from lookup so we know what's left to delete
            $existingTermsById.Remove($sourceId)

        }
        else {
            # ADD (Create New)
            if ($WhatIf) {
                Write-Log "  [WhatIf] Create Term: '$sourceName' ($sourceId)" -Color Green
                # Cannot recurse in WhatIf for new terms easily as parent doesn't exist
                Write-Log "    [WhatIf] (Would create children recursively...)" -Color DarkGray
            }
            else {
                Write-Log "  Creating Term: '$sourceName' ($sourceId)" -Color Green
                $newTerm = New-PnPTerm -TermSet $TermSetId -TermStore $TermStoreId -ParentTerm $ParentEntity.Id -Name $sourceName -Id $sourceId -Description $sourceDesc
                
                # Add Labels
                foreach ($label in $sourceLabels) {
                    if ($label -ne $sourceName) {
                        $newTerm.CreateLabel($label, 1033, $false)
                    }
                }
                Get-PnPContext | ForEach-Object { $_.ExecuteQuery() }

                # Recurse
                if ($sourceTerm.Terms) {
                    Sync-Terms -SourceTerms $sourceTerm.Terms -ParentEntity $newTerm -TermStoreId $TermStoreId -TermSetId $TermSetId
                }
            }
        }
    }

    # 2. Process Deletions (Terms in Target but not in Source)
    foreach ($termId in $existingTermsById.Keys) {
        $termToDelete = $existingTermsById[$termId]
        if ($WhatIf) {
            Write-Log "  [WhatIf] DELETE Term: '$($termToDelete.Name)' ($termId)" -Color Red
        }
        else {
            Write-Log "  DELETING Term: '$($termToDelete.Name)' ($termId)" -Color Red
            Remove-PnPTerm -Identity $termToDelete.Id -TermStore $TermStoreId -Force
        }
    }
}

# Main Execution
try {
    Write-Log "Reading Input File: $InputFile"
    Write-Log "Reading Input File: $InputFile"
    $data = Import-TermSetData -Path $InputFile

    Write-Log "Connecting to SharePoint: $SiteUrl"
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -ClientSecret $ClientSecret

    $termStore = Get-PnPTermStore
    if (-not $termStore) {
        Write-Error "Could not access Term Store."
        exit
    }
    
    # Hardcoded Names
    $TermGroupName = "Enterprise Term Sets"
    $TermSetName = "Organizational Hierarchy"

    # Get Term Group
    Write-Log "Looking for Term Group: $TermGroupName"
    $group = Get-PnPTermGroup -Identity $TermGroupName -TermStore $termStore.Id -ErrorAction SilentlyContinue
    if (-not $group) {
        Write-Error "Term Group '$TermGroupName' does not exist. Aborting."
        return
    }

    # Get Term Set
    Write-Log "Looking for Term Set: $TermSetName"
    $termSet = Get-PnPTermSet -Identity $TermSetName -TermGroup $group.Name -TermStore $termStore.Id -ErrorAction SilentlyContinue
    
    if (-not $termSet) {
        Write-Error "Term Set '$TermSetName' does not exist in Group '$TermGroupName'. Aborting."
        return
    }

    # Find the corresponding data in the input file
    $sourceSetData = $data.TermSets | Where-Object { $_.Name -eq $TermSetName }
    
    if (-not $sourceSetData) {
        Write-Warning "Could not find Term Set '$TermSetName' in the input file. Checking if there is only one set..."
        if ($data.TermSets.Count -eq 1) {
            $sourceSetData = $data.TermSets[0]
            Write-Log "Using the single Term Set found in input: '$($sourceSetData.Name)'" -Color Yellow
        }
        else {
            Write-Error "Input file does not contain a Term Set named '$TermSetName' and has multiple sets. Aborting."
            return
        }
    }

    # Sync Terms
    Write-Log "Syncing Terms for '$TermSetName'..." -Color Cyan
    if ($sourceSetData.Terms) {
        Sync-Terms -SourceTerms $sourceSetData.Terms -ParentEntity $termSet -TermStoreId $termStore.Id -TermSetId $termSet.Id
    }
    else {
        Write-Log "Source Term Set has no terms. Clearing target..." -Color Yellow
        Sync-Terms -SourceTerms @() -ParentEntity $termSet -TermStoreId $termStore.Id -TermSetId $termSet.Id
    }

    Write-Log "Sync Complete." -Color Green

}
catch {
    Write-Error "An error occurred: $_"
    Write-Error $_.ScriptStackTrace
}
