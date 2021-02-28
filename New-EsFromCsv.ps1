<#
.Synopsis
   Script for creating ES groups in AD from pre-recreated CSV file
.AUTHOR
   Yurii Baranchuk 13.04.20 
   Oleksandr Sobakar 21.12.20
.VERSION
   1.0 - initial
   2.0 - switch 'Local' added (if present, script creates Local AD group, otherwise - Global)
       - switch 'Dump' added (if present, script only creates csv file from existent AD groups)
.DESCRIPTION
   Script creates ES groups in AD in selected OU from selecting CSV file
   Required parameters:
   $File - CSV file location # EXAMPLE "C:\Temp\ES.csv"
   $Path - OU path location # EXAMPLE "OU=Endservices,OU=Groups,OU=LAB9,OU=Companies,DC=lab9,DC=local"
   Csv File EXAMPLE:
   Description,Name
   7Zip,ES_202249
   ........
.EXAMPLE
   .\Create-ESgroup-from-CSV.ps1 -File ".\ES.csv" -Path "OU=Endservices,OU=Groups,OU=LAB9,OU=Companies,DC=lab9,DC=local"
   .\Create-ESgroup-from-CSV.ps1 -File ".\ES.csv" -Path "OU=Endservices,OU=Groups,OU=LAB9,OU=Companies,DC=lab9,DC=local" -Local
   .\Create-ESgroup-from-CSV.ps1 -File ".\ES.csv" -Path "OU=Endservices,OU=Groups,OU=LAB9,OU=Companies,DC=lab9,DC=local" -Dump
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$File,

    [Parameter(Mandatory=$True,Position=2)]
    [string]$Path,

    [Parameter()]
    [switch]$Local,

    [Parameter()]
    [switch]$Dump

)

if ($Dump)
{
    Get-ADGroup -Filter * -SearchBase $Path -Properties Name,Description | `
        Select-Object -Property Name,Description | `
            Export-Csv -Path $File -NoTypeInformation
}
else
{
    Get-Content $File | ConvertFrom-Csv| ForEach-Object {
    if ($Local)
    {
        New-ADGroup -Name $_.Name `
                -Path $Path `
                -GroupScope DomainLocal `
                -GroupCategory Security `
                -Description $_.Description `
                -Verbose
    }
    else
    {
        New-ADGroup -Name $_.Name `
                -Path $Path `
                -GroupScope Global `
                -GroupCategory Security `
                -Description $_.Description `
                -Verbose    
    } 
  }
}
