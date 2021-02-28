function Get-Shortcuts{
<#
.SYNOPSIS
	Get properties of all shortcuts in StartMenu folder which name satisfy input pattern
.DESCRIPTION
	Get properties of all shortcuts in StartMenu folders (Current User and All Users ) which name satisfy input
	pattern. Pattern acceps wildcard

.EXAMPLE
	PS C:\> "7-Zip" | Get-Shortcuts
	Return properties for all shortcuts in StartMenu\7-Zip folder
.EXAMPLE
	PS C:\> Get-Shortcuts -Pattern "*Microsoft*"
	Return properties for all shortcuts in StartMenu\<AnyText>Microsoft<AnyText> folder 
.INPUTS Pattern
	String
.OUTPUTS PSObject
	Object include next properties
	Name - shortcut name
	Path - shortcut location
	Target - target path
	Args - target arguments
	WKDir - taget working directory
.NOTES
	General notes
#>
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory=$True,
            ValueFromPipeline=$true)]
        [string]$Pattern
    )
    BEGIN{
        $StartMenuDirs = 'CommonPrograms','Programs'
    } #BEGIN
    PROCESS{
        $list = $StartMenuDirs |
        ForEach-Object {
            Get-ChildItem "$([Environment]::GetFolderPath($PSItem))\*${Pattern}*" -Recurse
        } #ForEach-Object
    
        foreach ($shortcut in $list) {
            $WsShell = New-Object -ComObject "Wscript.Shell"
            $ShortcutObj = $WsShell.CreateShortcut($shortcut.FullName)
            $Props = @{
                Name = $shortcut.Name
                Path = $shortcut.FullName
                Target = $ShortcutObj.TargetPath
                Args = $ShortcutObj.Arguments
                WKDir = $ShortcutObj.WorkingDirectory
            } #$PROPS
            New-Object -TypeName PSObject -Property $props
        } #FOREACH
    } #PROCESS
    END{}
}