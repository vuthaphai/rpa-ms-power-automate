$objWord = New-Object -ComObject word.application
$objWord.Visible = $True
$objDoc = $objWord.Documents.Open("D:\00-RPA\Microsoft Power Automate\rpa-ms-power-automate\Project 4 - Sales Report Generation Process\Word_Files\Master_Template-David.docx.docx")
$objSelection = $objWord.Selection

function wordSearch($currentValue, $replaceValue) {
    $objSelection = $objWord.Selection
    $FindText = $currentValue
    $MatchCase = $false
    $MatchWholeWord = $true
    $MatchWildcards = $false
    $MatchSoundsLike = $false
    $MatchAllWordForms = $false
    $Forward = $true
    $wrap = $wdFindContinue
    $wdFindContinue = 1
    $Format = $false
    $ReplaceWith = $hash[$value]
    $ReplaceAll = 2

    $objSelection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $wrap, $Format, $ReplaceWith, $ReplaceAll)
}

$hash = @{"<Name>" = "%inputRow['Name']%"; "<Division>" = "%inputRow['Division']%"; "<Sales>" = "%inputRow['Sales']%" }

Write-Host "hash : " $hash.Keys
Write-Host "======================"

foreach ($value in $hash.Keys) {
    $currentValue = $value
    $replaceValue = $hash[$value]

    wordSearch $currentValue $replaceValue
    Write-Host "$currentValue : " $currentValue
    Write-Host "$replaceValue : " $replaceValue
    Write-Host "----------------------"

}
    
# Save the document to disk and close it. CHange $filename path to suit your environment.
$filename = "D:\00-RPA\Microsoft Power Automate\rpa-ms-power-automate\Project 4 - Sales Report Generation Process\Word_Files\Master_Template-David.docx.docx"
$objDoc.SaveAs([REF]$filename)
$objDoc.Close()

$objWord.Application.Quit()