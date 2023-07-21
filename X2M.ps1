$ExcelData = Get-Clipboard

$MarkdownData = (Write-Output $ExcelData | ForEach-Object {$_ `
    -replace "`t", "|" `
    -replace "^", "|"  `
    -replace "$", "|" `
    -replace "`n", "<br>" `
})

$ColNum = 0
for ($i=0; $i -lt ($MarkdownData[0]).Length; $i++ ){
    if ($MarkdownData[0][$i] -eq "|") {
        $ColNum++
    }
}
$ColNum--


$AlignConf = Read-Host "Set align Left:L, Center:C, Right:R"
switch ($AlignConf) {
    "L" { $AlignVar = ":---|" }
    "C" { $AlignVar = ":---:|" }
    "R" { $AlignVar = "---:|" }
    Default {$AlignVar = "---|"}
}
$Align ="|"
for ($i=0; $i -lt ([int]$ColNum); $i++){
    $Align +=$AlignVar     
}

# insert alignment into line 2 
$MarkdownData[1] =  $Align + "`n" + $MarkdownData[1]

# delete vertical bar in end
$MarkdownData = $MarkdownData[0..($MarkdownData.Length - 2)]
echo $MarkdownData
Set-Clipboard $MarkdownData