param (
    # Parameter help description
    [string]$Path = "D:\Project.INFP\Course.PowerShell"
)
$MSWordObject = New-Object -ComObject Word.Application
$docxFiles = Get-ChildItem $Path *.docx
New-Item -Path $Path -Name Examples -ItemType Directory -Force
foreach ($docxFile in $docxFiles){
    $FileName = $docxFile.BaseName
    $Code = $false
    $docxObject = $MSWordObject.Documents.Open($docxFile.FullName)
    $docxObject.Range().Paragraphs| ForEach-Object{
        if ($_.range().style.NameLocal -eq "Code") {
            $Code = $true
            $_.range().text.Replace('PS C:\> ','PS> ') | Out-File -FilePath "${Path}\Examples\${FileName}.txt" -Append
        } elseIf (($_.range().style.NameLocal -ne "Code") -and $Code) {
            '-'*120 | Out-File -FilePath "${Path}\Examples\${FileName}.txt" -Append
            $Code = $false
        }
    }
    $docxObject.Close()
}