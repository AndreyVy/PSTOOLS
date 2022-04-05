<#
.SYNOPSIS
Summary of what the script does
#>
[CmdletBinding()]
param(
   [Parameter(Mandatory = $false)]
   [string]$PathToOu='*',

   [Parameter(Mandatory = $false)]
   [string[]]$GpoName,

   # The name of the security principal for which to set the permission level
   [Parameter (Mandatory = $true)]
   [string]$IdentityName,

   # The type of security principal for which to set the permission level
   [Parameter (Mandatory = $false)]
   [ValidateSet('User', 'Group', 'Computer')]
   [string]$IdentityType = 'Group',

   # Specifies the permission level to set for the security principal
   [Parameter (Mandatory = $false)]
   [ValidateSet('GpoApply', 'GpoRead', 'GpoEdit', 'GpoEditDeleteModifySecurity', 'None')]
   [string]$PermissionLevel = 'GpoEdit'
)

# build search pattern for gpo accorfing to OU confition
if ($PathToOu -eq '*') { $SearchByOu = '.*' }
else {
	$SearchByGpoName = ( (get-gpinheritance -Target $PathToOu).gpolinks |
		ForEach-Object { "^$($_.DisplayName)$" } ) -join '|'
}

# build search pattern for gpo according to name condition
if ($null -eq $GpoName) {$GpoName = '.*'}
else {
	$SearchByGpoName = ( $GpoName | ForEach-Object {"^$_$" } ) -join '|'
}


Write-Verbose "Search pattern by OU: $SearchByOu"
Write-Verbose "Search pattern by name: $SearchByGpoName"

Get-GPO -All | Where-Object {$_.DisplayName -match $SearchByOu } |
	Where-Object {$_.DisplayName -match $SearchByGpoName } |
		Foreach-Object {
			$GpPermParams = @{
				'Guid'				= $_.Id
				'TargetName'		= $IdentityName
				'TargetType'		= $IdentityType
				'PermissionLevel'	= $PermissionLevel
				'Replace'			= $True
			}

			$null = Set-GPPermission @GpPermParams

			Write-Output "UserOutput: Setting permissions '$PermissionLevel' upon GPO '$($_.DisplayName)' for $IdentityType '$IdentityName'"
		}