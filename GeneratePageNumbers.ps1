$regexPageNumber = '<div class=.pageNumber.>'
$regexPageSection = '<div class=.footnote.[^>]*>'
$regexHeader1 = '^# (.*)$'
$regexHeader2 = "^#+ (.*) <div class='subsection'></div>"
$regexHeader3 = "^####+ ([a-zA-Z][a-zA-Z ]+)"
$regexNoMatch = "<!--No Match-->"
$header1 = ''
$header2 = ''
$header3 = ''
$index = 4
$Output = ""
$pageSectionFormattingPre = "<div class='footnote "
$pageSectionFormattingPre2 = "BG' style='"
$pageSectionFormattingPost = "text-align:center;font-weight:bold;height:20px;line-height:20px;font-size: 12px;width:656px;text-indent: 10px;color:white;`'>"
$sectionColorRed = "background: rgba(255,0,0,0.8);"
$sectionColorLime = "background: rgba(0,255,0,0.8);"
$sectionColorBlue = "background: rgba(0,0,255,0.8);"
$sectionColorYellow = "background: rgba(255,255,0,0.8);"
$sectionColorCyan = "background: rgba(0,255,255,0.8);"
$sectionColorMagenta = "background: rgba(255,0,255,0.8);"
$sectionColorSilver = "background: rgba(139, 139, 139, 0.8);"
$sectionColorMaroon = "background: rgba(128,0,0,0.8);"
$sectionColorOlive = "background: rgba(128,128,0,0.8);"
$sectionColorGreen = "background: rgba(0,128,0,0.8);"
$sectionColorPurple = "background: rgba(128,0,128,0.8);"
$sectionColorTeal = "background: rgba(0,128,128,0.8);"
$sectionColorNavy = "background: rgba(0,0,128,0.8);"

$sectionColorArray = @($sectionColorRed,$sectionColorLime,$sectionColorBlue,$sectionColorYellow,$sectionColorCyan,$sectionColorMagenta,$sectionColorSilver,$sectionColorMaroon,$sectionColorOlive,$sectionColorGreen,$sectionColorPurple,$sectionColorTeal,$sectionColorNavy) # 13 colors
$sectionColorIndex = 0

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

foreach($line in Get-Content $File)
{
    if($line -match $regexPageNumber)
    {
        $line=$matches[0] + ($index++) + "</div>"
    }

    if($line -match $regexHeader1)
    {
        $header1 = $matches[1]
		$header2 = ''
		$header3 = ''
		$sectionColorIndex++
		if($sectionColorIndex -ge $sectionColorArray.Length)
        {
			$sectionColorIndex = 0
        }
    }
	if($line -match $regexHeader2)
    {
		$header3 = ''
        $header2 = ' &rarr; ' + $matches[1]
		if($Selection.ToLower() -ne 'm')
		{
        	$sectionColorIndex++
			if($sectionColorIndex -ge $sectionColorArray.Length)
        	{
				$sectionColorIndex = 0
        	}
		}
		Write-Host $line
    }
	if($Selection.ToLower() -ne 'm')
	{
		if($line -match $regexHeader3 -and $header2 -ne "")
    	{
			$tempHeader3 = ' &rarr; ' + $matches[1]

			if($line -notmatch $regexNoMatch)
			{
        		$header3 = $tempHeader3
			}
    	}
	}
    if($line -match $regexPageSection)
    {
    	$line = $pageSectionFormattingPre + ($header1 -replace "[ \(\)&]", "") + $pageSectionFormattingPre2 + $sectionColorArray[$sectionColorIndex] +  $pageSectionFormattingPost + $header1 + $header2 + $header3 + "</div>"
    }
    $Output += $line + "`r`n"
}

set-content $File $Output.TrimEnd("`r`n");
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear();