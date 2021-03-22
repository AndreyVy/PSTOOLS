<#
.SYNOPSIS
    Converts RES export to PowerShell
.DESCRIPTION
    Creates one script per application (xml may contain several apps). Generates code for
        - Application start
        - Registry add
        - Environment variables add
.EXAMPLE
    .\Start.ps1 -FullName "C:\data\start_schlumberger_olga 2017.2.0_olga 2017.2.0.xml"
    Creates .\Output directory and save ps1 scripts there. 
.EXAMPLE
    dir C:\data *.xml | .\Start.ps1 -Verbose
    Creates .\Output directory and creates ps1 scripts for each applications from xml files in C:\data directory
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
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