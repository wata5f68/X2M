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

# insert alignment into line 2 
$MarkdownData[1] =  $align + "`n" + $MarkdownData[1]

# delete vertical bar in end
$MarkdownData = $MarkdownData[0..($MarkdownData.Length - 2)]
Set-Clipboard $MarkdownData