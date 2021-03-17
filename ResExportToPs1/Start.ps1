[CmdletBinding()]
param(
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [ValidateScript( { Test-Path $_ })]
    [string[]]$FullName
)
BEGIN {
    . $PSScriptRoot\functions.ps1
}
PROCESS {
    foreach ($file in $FullName) {
        Write-Verbose "Parse file $file"
        [xml]$BuildingBlock = Get-Content -Path $file
        $ResData = ConvertFrom-ResExport -XmlData $BuildingBlock
	Write-Debug "after convertfrom-resexport"
        foreach ($ResObj in $ResData) {
            New-ShortcutScript -ResObj $ResObj
        }       
    } #foreach
} #PROCESS
END {}