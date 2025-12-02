<#
.SYNOPSIS
    Extracts a SharePoint Term Set to a .psd1 file in a specific format.

.DESCRIPTION
    This script connects to SharePoint Online using PnP PowerShell, retrieves a specified Term Set,
    and exports its terms (Name, ID, Description, Labels, and nested Terms) into a .psd1 file.
    The output format mimics a nested hashtable structure.

.PARAMETER TermSetName
    The name of the Term Set to extract. Default is "Organizational Hierarchy".

.PARAMETER SiteUrl
    Optional. The URL of the SharePoint site to connect to. If not provided, assumes an active PnP connection.

.PARAMETER OutputFile
    The path to the output .psd1 file. Default is ".\TermSetExport.psd1".

.EXAMPLE
    .\Extract-TermSet.ps1 -TermSetName "Organizational Hierarchy" -SiteUrl "https://contoso.sharepoint.com/sites/intranet" -OutputFile "C:\Temp\OrgHierarchy.psd1"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$TermSetName = "Organizational Hierarchy",

    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = ".\TermSetExport.psd1"
)

# Function to format a single term recursively
function Get-TermObject {
    param (
        [Microsoft.SharePoint.Client.Taxonomy.Term]$Term,
        [int]$IndentLevel = 0
    )

    $indent = " " * ($IndentLevel * 4)
    $innerIndent = " " * (($IndentLevel + 1) * 4)
    
    # Get basic properties
    $termName = $Term.Name
    $termId = $Term.Id.ToString()
    $termDesc = if ($Term.Description) { $Term.Description } else { "" }
    
    # Get Labels (Synonyms) - excluding the default label which is the Name
    # Note: PnP/CSOM might return all labels. We filter for non-default or just dump all if that's the requirement.
    # Looking at the sample, "Labels" seems to contain the Name (e.g. "CMMN") and the Description/Full Name (e.g. "Communities")?
    # Or maybe it's just a list of labels.
    # In the sample: Name = "CMMN", Labels = @("CMMN", "Communities")
    # So it includes the name itself.
    $labels = @()
    if ($Term.Labels) {
        $labels = $Term.Labels | ForEach-Object { $_.Value }
    }
    
    # Get Child Terms
    # We need to ensure we load them
    $childTerms = $Term.Terms
    # If context isn't loaded, we might need to load them, but PnP usually handles this if we iterate.
    # However, for deep recursion, it's safer to ensure they are available.
    # Assuming PnP Get-PnPTerm with -Includes Terms works or we iterate.
    
    # Construct the string representation directly to ensure formatting
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("$indent@{") | Out-Null
    $sb.AppendLine("${innerIndent}Name = `"$termName`";") | Out-Null
    $sb.AppendLine("${innerIndent}ID = `"$termId`";") | Out-Null
    
    if (-not [string]::IsNullOrEmpty($termDesc)) {
        $sb.AppendLine("${innerIndent}Description = `"$termDesc`";") | Out-Null
    }
    
    # Process Child Terms
    if ($childTerms.Count -gt 0) {
        $sb.AppendLine("${innerIndent}Terms = @(") | Out-Null
        $count = 0
        foreach ($child in $childTerms) {
            $childStr = Get-TermObject -Term $child -IndentLevel ($IndentLevel + 1)
            # Remove the last newline from childStr to handle commas correctly if needed, 
            # but psd1 arrays usually just need items separated by newlines or commas.
            # The sample uses commas? Let's check the sample.
            # Sample:
            # Terms = @(
            #    @{ ... },
            #    @{ ... }
            # );
            # So yes, commas between objects.
            
            $sb.Append($childStr) | Out-Null
            $count++
            if ($count -lt $childTerms.Count) {
                $sb.AppendLine(",") | Out-Null
            } else {
                $sb.AppendLine("") | Out-Null
            }
        }
        $sb.AppendLine("${innerIndent});") | Out-Null
    }

    # Process Labels
    if ($labels.Count -gt 0) {
        $labelStr = $labels | ForEach-Object { "`"$_`"" }
        $labelJoined = $labelStr -join ", "
        $sb.AppendLine("${innerIndent}Labels = @($labelJoined);") | Out-Null
    }

    $sb.Append("$indent}")
    return $sb.ToString()
}

# Main Execution
try {
    if ($SiteUrl) {
        Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
        Connect-PnPOnline -Url $SiteUrl -Interactive
    }

    Write-Host "Retrieving Term Set: $TermSetName..." -ForegroundColor Cyan
    # Get the term set and include all terms recursively
    # Note: Get-PnPTermSet returns the set. To get terms recursively, we might need Get-PnPTerm.
    # But Get-PnPTerm -TermSet ... -Recursive returns a flat list.
    # We want the hierarchy.
    # The TermSet object has a "Terms" property which are the root terms.
    # We need to make sure we load the full tree.
    
    $termSet = Get-PnPTermSet -Identity $TermSetName -Includes Terms, Terms.Terms, Terms.Terms.Terms, Terms.Terms.Terms.Terms, Terms.Terms.Terms.Terms.Terms
    # Note: The Includes depth might be limited. A better way is to load recursively or use the flat list and build the tree.
    # But for a script, let's try to traverse.
    # If the tree is very deep, explicit includes might fail.
    # Alternative: Get all terms flat, then build hierarchy?
    # Or just traverse and load on demand (slower but works).
    # Let's assume standard PnP context loading.
    
    if (-not $termSet) {
        Write-Error "Term Set '$TermSetName' not found."
        exit
    }

    $rootTerms = $termSet.Terms
    $context = Get-PnPContext
    $context.Load($rootTerms)
    $context.ExecuteQuery()
    
    # We need a recursive loader because the Includes above is messy and fragile.
    function Load-TermsRecursively {
        param ($terms)
        
        $ctx = Get-PnPContext
        $ctx.Load($terms)
        $ctx.ExecuteQuery()
        
        foreach ($term in $terms) {
            $ctx.Load($term.Terms)
            $ctx.Load($term.Labels)
            $ctx.ExecuteQuery()
            
            if ($term.Terms.Count -gt 0) {
                Load-TermsRecursively -terms $term.Terms
            }
        }
    }
    
    Write-Host "Loading terms (this may take a moment)..." -ForegroundColor Cyan
    Load-TermsRecursively -terms $rootTerms

    # Start building the output
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("@{") | Out-Null
    $sb.AppendLine("     TermSets = @(") | Out-Null
    $sb.AppendLine("        @{") | Out-Null
    $sb.AppendLine("             Name = `"$($termSet.Name)`";") | Out-Null
    $sb.AppendLine("             ID = `"$($termSet.Id)`";") | Out-Null
    $sb.AppendLine("             Description = `"$($termSet.Description)`";") | Out-Null
    $sb.AppendLine("             Terms = @(") | Out-Null

    $count = 0
    foreach ($term in $rootTerms) {
        $termStr = Get-TermObject -Term $term -IndentLevel 4
        $sb.Append($termStr) | Out-Null
        $count++
        if ($count -lt $rootTerms.Count) {
            $sb.AppendLine(",") | Out-Null
        } else {
            $sb.AppendLine("") | Out-Null
        }
    }

    $sb.AppendLine("             );") | Out-Null
    $sb.AppendLine("        }") | Out-Null
    $sb.AppendLine("     )") | Out-Null
    $sb.AppendLine("}") | Out-Null

    # Write to file
    $sb.ToString() | Out-File -FilePath $OutputFile -Encoding UTF8
    Write-Host "Export complete. File saved to: $OutputFile" -ForegroundColor Green

} catch {
    Write-Error "An error occurred: $_"
}
