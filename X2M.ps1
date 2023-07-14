$ExcelData = Get-Clipboard  
$ColNum = ($ExcelData[0] | Measure-Object -Word).Words

$MarkdownData = (Write-Output $ExcelData | ForEach-Object {$_ `
    -replace "`t", "|" `
    -replace "^", "|"  `
    -replace "$", "|" `
    -replace "`n", "<br>" `
})

$align ="|"
for ($i=0; $i -lt ([int]$ColNum); $i++){
    $align +="---|"     
}

$MarkdownData[1] =  $align + "`n" + $MarkdownData[1]
$MarkdownData = $MarkdownData[0..($MarkdownData.Length - 2)]
Set-Clipboard $MarkdownData