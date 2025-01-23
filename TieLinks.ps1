# Define the input and output file paths
$inputFile = ""
$outputFile = "test.txt"

$FileBiota = "D:\The Wandering Planet\Scripts\WPBiota.txt"
$FilePlayer = "D:\The Wandering Planet\Scripts\WPPlayer.txt"
$FileTravel = "D:\The Wandering Planet\Scripts\WPTravel.txt"
$FileVehicles = "D:\The Wandering Planet\Scripts\WPVehicles.txt"
$FileMutation = "D:\The Wandering Planet\Scripts\WPMutation.txt"

$Selection = Read-Host "(B)iota, (P)layer, (M)utation, (T)ravel, (V)ehicles"

$File = ""

switch ($Selection.ToLower())
{
    "b" {
		$File = $FileBiota
		Break
	}
	"p" {
		$File = $FilePlayer
		Break
	}
	"m" {
		$File = $FileMutation
		Break
	}
	"t" {
		$File = $FileTravel
		Break
	}
	"v" {
		$File = $FileVehicles
		Break
	}
}

$inputFile = $File
$outputFile = $File

# Read the content of the input file
$content = Get-Content -Path $inputFile -Raw

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
Set-Content -Path $outputFile -Value $content

Write-Host "Processing complete. Updated file saved to $outputFile."

$Selection = Read-Host "Exit?"