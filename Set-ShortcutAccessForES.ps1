<#
.Synopsis
   The script adds AD groups to shortcuts from pre-recreated CSV file
.AUTHOR
   Yurii Baranchuk 05.05.20
   Oleksandr Sobakar 18.06.20
   Andrii Hadzevych 23.06.20
   Oleksandr Sobakar 10.11.20
.VERSION
   1.0 - initial
   1.1 - disable inhiritance and remove company group (ITxx)
   1.2 - added -dump switch
   1.3 - remove Authenticated Users group
.DESCRIPTION
   The script adds AD groups to shortcuts from selected CSV file
   Required parameters:
   $File - CSV file location # EXAMPLE "C:\Temp\ES.csv"
   $Path - location of shortcuts # EXAMPLE "\\Lab9.local\root\Commonsystem\Startmenu\"
   Optional parameters:
   $dump - create CSV file from the existing shortcuts
   Csv File EXAMPLE:
   ADDescription,ADName
   7Zip,ES_202249
   ........
.EXAMPLE
   .\Add-ESgroup-to-shortcut.ps1 -File "C:\Temp\ES.csv" -Path "\\Lab9.local\root\Commonsystem\Startmenu\"
#>
[CmdletBinding()]
Param(
	[Parameter(Mandatory = $True, Position = 1)]
	[string]$File,

	[Parameter(Mandatory = $True, Position = 2)]
	[string]$Path,
	
	[Parameter(Mandatory = $false)]
	[switch]$dump

)

if ($dump){
	Get-ChildItem -Path $Path -Filter '*.lnk' -File -Recurse -Force | select-object @{n="ADDescription";e={$_.baseName}},@{n="ADName";e={((((get-acl $_.fullName).access.identityReference.value -like "*\ES_*") + @(".\ES_000000"))[0].split("\"))[1]}} | where-object {$_.ADName -ne "ES_000000"} | export-csv -path $file -noTypeInformation
	exit
}

$groups = Import-Csv -Path $file

# get a list of *.lnk FileIfo objects where the file's BaseName can be found in the
# CSV column 'Name'. Group these files on their BaseName properties
$linkfiles = Get-ChildItem -Path $Path -Filter '*.lnk' -File -Recurse -Force |
Where-Object { $groups.ADDescription -contains $_.BaseName } |
Group-Object BaseName

# iterate through the grouped *.lnk files
$linkfiles | ForEach-Object {
	$baseADDescription = $_.Name  # the name of the Group is the BaseName of the files in it
	$adGroup = ($groups | Where-Object { $_.ADDescription -eq $baseADDescription }).ADName
    
	# create a new access rule
	$rule = [System.Security.AccessControl.FileSystemAccessRule]::new($adGroup, "ReadAndExecute", "Allow")

	$_.Group | ForEach-Object {
		# get the current ACL of the file
		$acl = Get-Acl -Path $_.FullName
		# disable inhiritance and apply
		$acl.SetAccessRuleProtection($true, $true)
		(Get-Item $_.FullName).SetAccessControl($acl)# use this method to avoid executing script with elevated privileges
		# remove company group (ITxx) from the ACL
		$acl = Get-Acl -Path $_.FullName
		$ruleIT = $acl.access | Where-Object { $_.IdentityReference -eq "$env:USERDOMAIN\$env:USERDOMAIN" }
        
		if ($ruleIT) {
			$acl.RemoveAccessRule($ruleIT) | Out-Null
		}
        # remove Authenticated Users group from the ACL
        $ruleAuth = $acl.access | Where-Object { $_.IdentityReference -eq "NT AUTHORITY\Authenticated Users" }
        if ($ruleAuth) {
			$acl.RemoveAccessRule($ruleAuth) | Out-Null
		}
		# add the new rule to the ACL
		$acl.SetAccessRule($rule)
		(Get-Item $_.FullName).SetAccessControl($acl)# use this method to avoid executing script with elevated privileges
		# output for logging csv
		[PsCustomObject]@{
			'Group' = $adGroup
			'File'  = $_.FullName
		}
	}
 
}

function Update-ShortcutAfterImport {
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
}