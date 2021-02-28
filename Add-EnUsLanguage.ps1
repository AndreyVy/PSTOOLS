Param(
    [string]$LangID = "en-US"
)

$ll = Get-WinUserLanguageList
if (-not ($ll.Exists({$args[0].LanguageTag -eq $LangID}))){
$ll.Clear() 
$ll.Insert(0,$LangID)
Set-WinUserLanguageList -LanguageList $ll -Force
}