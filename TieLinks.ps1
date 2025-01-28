# Define the input and output file paths
$File = Read-Host "Enter Path to File:"

# Read the content of the input file
$content = Get-Content -Path $File -Raw

# Define regex patterns
$pageNumberPattern = "<div class='pageNumber'> *(\d+) *</div>"
$footnotePattern = "<div class='footnote[^']*' style='[^>]+'>[^<]*; ?([^<]+)</div>"
$linkPattern = "\[{{ *([^}]+) *}}{{}}\]\(#p(\d+)\)"

# Create a hash table to store page numbers and titles
$pageTitleMap = @{}

# Extract page numbers and titles and populate the hash table
$matches = [regex]::Matches($content, "(?s)$pageNumberPattern.*?$footnotePattern")
foreach ($match in $matches) {
    $pageNumberMatch = [regex]::Match($match.Value, $pageNumberPattern)
    $footnoteMatch = [regex]::Match($match.Value.Trim(), $footnotePattern)

    if ($pageNumberMatch.Success -and $footnoteMatch.Success) {
        $pageNumber = $pageNumberMatch.Groups[1].Value.Trim()
        $title = $footnoteMatch.Groups[1].Value.Trim()
        $pageTitleMap[$title] = $pageNumber

        # Output matches to the console
        Write-Host "Match found: Title = $title, PageNumber = $pageNumber"
    } else {
        # Output unmatched cases for debugging
        Write-Host "No match found for block: $($match.Value)"
    }
}

# Process the content and replace placeholders
$content = [regex]::Replace($content, $linkPattern, {
    param($match)
    $titleKey = $match.Groups[1].Value.Trim()
    $pageKey = $match.Groups[2].Value.Trim()
    if ($pageTitleMap.ContainsKey($titleKey)) {
        $replacement = "[{{ " + $titleKey + " }}{{}}](#p" + $pageTitleMap[$titleKey] + ")"
        Write-Host "Replaced link: Original = $($match.Value), Replacement = $replacement"
        $replacement
    } else {
        Write-Host "No replacement found for link: $($match.Value)"
        $match.Value  # Keep the original if no match is found
    }
})

# Save the updated content to the output file
Set-Content -Path $File -Value $content

Write-Host "Processing complete. Updated file saved to $File."

$Selection = Read-Host "Exit?"