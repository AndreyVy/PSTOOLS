<#
.SYNOPSIS
    Converts RES export to PowerShell
.DESCRIPTION
    Creates one script per application (xml may contain several apps). Generates code for
        - Application start
        - Registry
        - Environment variables
        - Scripts
        - Linked actions
    By default, output is saved to <CURRENT CONSOLE WORKING DIRECTORY>\Scripts.
.EXAMPLE
    PS C:\Temp> C:\Scripts\Convert-ResToScript.ps1 -Path "C:\data\start_schlumberger_olga 2017.2.0_olga 2017.2.0.xml"
    Creates C:\Temp\Scripts directory and save ps1 scripts there. If C:\Temp\Scripts already has ps1-files with app
    names, script will generate error.
.EXAMPLE
    PS C:\Temp> C:\Scripts\Convert-ResToScript.ps1 -Path "C:\data\start_schlumberger_olga 2017.2.0_olga 2017.2.0.xml" -Force
    Do the same as previous example.
    If C:\Temp\Scripts already contains scripts from previous launches, they will be overwritten
.EXAMPLE
    dir C:\data *.xml | .\Convert-ResToScript.ps1 -OutputDir C:\Results -Verbose
    Creates C:\Results directory and save ps1 scripts for each applications described in xml files from C:\data directory
    Console will contain verbose information 
.INPUTS
    XML file exported from RES
.OUTPUTS
    PowerShell script
.NOTES
    Version 2.5
        * Recognize start of another RES application
    Version 2.4
        * Info about RES shortcut permissions moved on top of the script. Example:
                # user GCLOUD\boa
                # group GCLOUD\ES_202548
                # user GCLOUD\testuser
                function Start-Application {
        * Add assigned permissions on registry/variables/actions. Example:
                # Workspace: AE
                # GROUP GCLOUD\AE
                # USER GCLOUD\boa
                # notingroup GROUP GCLOUD\AE_YE
                $EnvName  = 'SLBSLS_LICENSE_FILE'
        * Add powerzone information.Example:
                # Name: Win 10 TaskBar
                # GLOBAL
                # POWERZONE Windows 10
                (Get-Item HKCU:).CreateSubKey("HKEY_CURRENT_USER\",$true) | Out-Null
        * fix registry with empty ""
        * Refactor code to make it more clear
        * Small fixes in output script formatting.
    Version 2.3
        * fix bug for Path parameter with Resolve-Path
        * Resolve dos-style %variables%
        * Add information about disabled shortcuts: by Disable option, by StartMenu item option
        * Add partial information about access groups and users
        * Add actions on app close
        * Change info appearance: additional comments will be added to result script if requested info has
          non-default values. For example, if workspace was not set, result script will not have workspace
          comment at all. If workspace has specific value, script will have workspace comment with value
    Version 2.2
        * Script adjusted to work with logon actions
        * removed undesired blank lines
        * added additional info about to each res object (name or description)
        *changed default output dir to current working directory. See examples
    TO-DO:
    - Folder Redirections
    - mapping
    - embeddedpolicies
#>
[CmdletBinding()]
param(
    [Parameter( Mandatory = $true, ValueFromPipeline = $true)]
    $Path,
    [string]$OutputDir = "$PWD\Scripts",
    [switch]$PassThru,
    [switch]$Force
)
BEGIN{
$ScriptTemplate = @'
<INFO>function Start-Application {
<LINKEDAPPS><ENVVAR><REG_ENTRY><SCRIPTS><STARTAPP>
}
Start-Application
'@

$LinkedScriptTemplate = @'
<INFO><LINKEDAPPS><ENVVAR><REG_ENTRY><SCRIPTS>
'@

$StartApp = @'
    Start-Process -FilePath "<PATHTOAPP>"<ARGS><WKDIR> -WindowStyle Normal
'@

function ConvertResExport {
    [CmdletBinding()]
    param([xml]$XmlData)
    $applications = $XmlData.respowerfuse.buildingblock.application
        
    # logon actions do not have application tag
    if ($null -eq $applications) {
        $applications = $XmlData.respowerfuse.buildingblock
    }
        
    foreach ($app in $applications) {
            
        $EnvVariables = [System.Collections.ArrayList]@()
        foreach ($variable in $app.powerlaunch.variable) {
            $null = $EnvVariables.Add(
                [PsCustomObject]@{
                    'Name'       = "$($variable.name)"
                    'Value'      = "$($variable.value)"
                    'Enabled'    = "$($variable.enabled)"
                    'Workspace'  = $Workspaces | Foreach-Object { if ($_.Guid -in $variable.workspacecontrol.workspace) { $_.Name } }
                    'AccessInfo' = GetAccessControls -Target $variable
                })   #   save as hashtable
        } # foreach $variable
        
        $ResScripts = [System.Collections.ArrayList]@()
        foreach ($script in $app.powerlaunch.exttask) {
            $null = $ResScripts.Add(
                [PsCustomObject]@{
                    'Name'       = $script.description
                    'Command'    = $script.Command
                    'Text'       = $script.script
                    'Type'       = $script.scriptext
                    'Enabled'    = $script.enabled
                    'Workspace'  = $Workspaces | Foreach-Object { if ($_.Guid -in $script.workspacecontrol.workspace) { $_.Name } }
                    'AccessInfo' = GetAccessControls -Target $script
                })     #   save as object
        } # foreach $script

        $OnCloseScripts = [System.Collections.ArrayList]@()
        foreach ($script in $app.powerlaunch.exttaskex) {
            $null = $OnCloseScripts.Add(
                [PsCustomObject]@{
                    'Name'       = $script.description
                    'Command'    = $script.Command
                    'Text'       = $script.script
                    'Type'       = $script.scriptext
                    'Enabled'    = $script.enabled
                    'Workspace'  = $Workspaces | Foreach-Object { if ($_.Guid -in $script.workspacecontrol.workspace) { $_.Name } }
                    'AccessInfo' = GetAccessControls -Target $script
                })     #   save as object
        } # foreach $script

        $ResRegistry = [System.Collections.ArrayList]@()
        foreach ($regEntry in $app.powerlaunch.registry) {
            $null = $ResRegistry.Add(
                [PsCustomObject]@{
                    'Name'       = $regEntry.name
                    'regText'    = [System.Text.Encoding]::ASCII.GetString( [byte[]] -split ($regEntry.registryfile -replace '..', '0x$& ') )
                    'Enabled'    = $regEntry.enabled
                    'Workspace'  = $Workspaces | Foreach-Object { if ($_.Guid -in $regEntry.workspacecontrol.workspace) { $_.Name } }
                    'AccessInfo' = GetAccessControls -Target $regEntry
                }
            )
        } # foreach $regEntry

        $fta = [System.Collections.ArrayList]@()
        foreach ($extension in $app.instantfileassociations.association) {
            $null = $fta.Add(
                @{
                    'extension'       = $extension.extension
                    'command'         = $extension.command
                    'parameters'      = $extension.parameters
                    'description'     = $extension.description
                    'dde_enabled'     = $extension.dde_enabled
                    'dde_message'     = $extension.dde_message
                    'dde_application' = $extension.dde_application
                    'dde_topic'       = $extension.dde_topic
                    'dde_ifexec'      = $extension.dde_ifexec
                } #hashtable fta
            ) # add
        } #foreach extension

        $LinkedApplications = [System.Collections.ArrayList]@()
        $LinkedActions = $app.powerlaunch.linked_actions
        foreach ($LinkedActionGuid in $LinkedActions) {
            $LinkedApplication = ( $applications | Where-Object { $_.guid -eq $LinkedActionGuid.linked_to_application } ).configuration.title
            if ($LinkedApplication) {
                $null = $LinkedApplications.Add($LinkedApplication)
            }
        } # foreach LinkedActionGuid

        $Target = $app.configuration.commandline
        $Arguments = $app.configuration.parameters
        $WorkingDir = $app.configuration.workingdir

        # recognize start of another RES application
        if ($Target -eq '%respfdir%\pwrgate.exe') {
            
            Write-Verbose 'RES <commandline> setting has pwrgate.exe as target. The shortcuts call for another RES application'
            
            $ResAppId = $Arguments -split ' ' | Select-Object -First 1
            if ($ResAppId -notmatch '[^0-9]') {
                
                Write-Warning "Called application has id $ResAppId in RES database"
                
                $Target = ($Arguments -split ' ' | Select-Object -Skip 1) -join ''
                $Arguments = ''
                $WorkingDir = ''
            }
        }

        [PsCustomObject][ordered]@{
            'Name'         = $app.configuration.title
            'Description'  = $app.configuration.description
            'Target'       = $Target
            'Arguments'    = $Arguments
            'WorkingDir'   = $WorkingDir
            'Scripts'      = $ResScripts
            'CloseScripts' = $OnCloseScripts
            'Variables'    = $EnvVariables
            'Registry'     = $ResRegistry
            'FTA'          = $fta
            'IsShortcut'   = ( '-' -ne "$($app.configuration.commandline)" ) -and ( $null -ne $($app.configuration) )
            'ESNumber'     = $app.accesscontrol.grouplist.group.InnerText
            'LinkedApps'   = $LinkedApplications
            'InStartMenu'  = $app.configuration.createmenushortcut
            'IsEnabled'    = $app.settings.enabled
            'AccessInfo'   = GetAccessControls -Target $app
        } #hashtable output
    } # foreach $app
} # function ConvertResExport

function ConvertRegToScript {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$RegData
    )
    $RegistryTypes = @{ "string" = "String";
        "hex"                    = "Binary";
        "dword"                  = "DWord";
        "hex(b)"                 = "QWord";
        "hex(7)"                 = "MultiString";
        "hex(2)"                 = "ExpandString";
        "hex(0)"                 = "Unknown"
    }
    $RegistryHives = @{ "HKEY_LOCAL_MACHINE\\" = "HKLM:";
        "HKEY_CURRENT_USER\\"                  = "HKCU:";
        "HKEY_CLASSES_ROOT\\"                  = "HKCR:";
        "HKEY_USERS\\"                         = "HKU:";
        "HKEY_CURRENT_CONFIG\\"                = "HKCC:"
    }
    $key = ""
    $name = ""
    $value = ""
    $type = ""
    $curvalue = ""
    $code = [System.Collections.ArrayList]@()

    New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR | Out-Null
    New-PSDrive -PSProvider registry -Root HKEY_USERS -Name HKU | Out-Null
    New-PSDrive -PSProvider registry -Root HKEY_CURRENT_CONFIG -Name HKCC | Out-Null

    $RegStrings = $RegData -split "`r`n"
    foreach ($line in $RegStrings) {
        if ($line.Length -gt 0) {
            # current line has comment
            if ($line -match '(^;)|(^Windows Registry Editor)|(^""$)' ) {
                Continue;
            }
            # currrent line is a registry key            
            if ($line -match '^\[([-]*)(.+)\]$' ) {
                $key = $matches[2]
                $remove = $matches[1]
                foreach ($item in $RegistryHives.Keys) {
                    if ($key -match $item) {
                        $parentKey = $RegistryHives[$item]
                        $subKey = $key -replace "($item)(.)", '$2'
                        $key = $parentKey + "\" + $subKey
                    }
                }
                # if registry key should be removed                    
                if ($remove -eq '-') {
                    $null = $code.add("(Get-Item $parentKey).DeleteSubKeyTree(`"$subKey`")")
                }
                else {
                    $null = $code.add("(Get-Item $($parentKey)).CreateSubKey(`"$($subKey)`",`$true) | Out-Null")
                }
                Continue;
            }
            # current line is a pair of registry name and value
            if ($line -match '(".+"|@)(\s*=\s*)(.+)') {
                $subline1 = $Matches[1]
                $subline2 = $Matches[3]
                $name = $($subline1.Trim() -replace '(^")|("$)', '') -replace "([\\])(.)", '$2'
                if ($name -eq "@") { $name = "" }
                if ($line -match "[\\]$") {
                    $curvalue += $subline2
                    Continue
                }
                else { $value = $Matches[3] }
            }
            else {
                # current registry value continues on the next line
                if ($line -match "[\\]$") {
                    if (![string]::IsNullOrEmpty($curvalue)) {
                        $curvalue += $line
                        Continue
                    }
                }
                else {
                    if (![string]::IsNullOrEmpty($curvalue)) {
                        $value = $curvalue + $line
                        $curvalue = ""
                    }
                }
            }
            # current registry name should be removed
            if ($value -eq "-") {
                $code += "(Get-Item $parentKey).OpenSubKey(`"$subKey`",`$true).DeleteValue(`"$name`")"
                Continue;
            }
            # parsing registry type
            elseif ($value -match "^((hex:)|(hex\(0\):)|(hex\(2\):)|(hex\(7\):)|(hex\(b\):)|(dword:))") {
                $type, $value = $value -split ":", 2
                $type = $RegistryTypes[$type]
                if ($value -match "[\\\s +]") {
                    $value = $value -replace "[\\\s +]", ""
                }
            }
            elseif ( ($value -match '(^")') -and ($value -match '("$)')) {
                $type = "String"
                $value = $($value.Trim() -replace '(^")|("$)', '') -replace "([\\])(.)", '$2'
            }
            # processing values according to their types
            switch ($type) {
                "String" { $value = $value -replace '"', '""' }
                "Binary" { $value = $value -split "," | ForEach-Object { [System.Convert]::ToInt64($_, 16) } }
                "QWord" {
                    $temparray = @()
                    $temparray = $value -split ","
                    [array]::Reverse($temparray)
                    $value = -join $temparray
                }
                "MultiString" {
                    $MultiStrings = [System.Collections.ArrayList]@()
                    $temparray = @()
                    $temparray = $value -split ",00,00,00,"
                    for ($i = 0; $i -lt ($temparray.Count - 1); $i++) { 
                        $val = ([System.Text.Encoding]::Unicode.GetString((($temparray[$i] -split ",") + "00" | ForEach-Object { [System.Convert]::ToInt64($_, 16) }))) -replace '"', '""'
                        $null = $MultiStrings.Add( $val )
                    }
                    $value = $MultiStrings
                }
                "ExpandString" {
                    if ($value -match "^00,00$") { $value = "" }
                    else {
                        $value = $value -replace ",00,00$", ""
                        $value = [System.Text.Encoding]::Unicode.GetString((($value -split ",") | ForEach-Object { [System.Convert]::ToInt64($_, 16) }))
                        $value = $value -replace '"', '""'
                    }
                }
                "Unknown" {
                    $null = $code.Add("# Unknown registry type is not supported [$key]\$name, type `"$type`"")
                    Continue;                
                }
                     
            }
            $name = $name -replace '"', '""'
            if (($type -eq "Binary") -or ($type -eq "Unknown")) {
                $value = "@(" + ($value -join ",") + ")"
                $null = $code.add("(Get-Item $($parentKey)).OpenSubKey(`"$($subKey)`",`$true).SetValue(`"$($name)`",[byte[]] $($value),`"$($type)`")")
            }
            elseif ($type -eq "MultiString") {
                $value = "@(" + ('"' + ($value -join '","') + '"') + ")"
                $null = $code.add("(Get-Item $($parentKey)).OpenSubKey(`"$($subKey)`",`$true).SetValue(`"$($name)`",[string[]] $($value),`"$($type)`")")
            }    
            elseif (($type -eq "DWord") -or ($type -eq "QWord")) {
                $value = "0x" + $value
                $null = $code.add("(Get-Item $($parentKey)).OpenSubKey(`"$($subKey)`",`$true).SetValue(`"$($name)`",$($value),`"$($type)`")")
            }
            else {
                $null = $code.add("(Get-Item $($parentKey)).OpenSubKey(`"$($subKey)`",`$true).SetValue(`"$($name)`",`"$($value)`",`"$($type)`")")
            }
        }
    }
    $code
    Remove-PSDrive -Name HKCR
    Remove-PSDrive -Name HKU
    Remove-PSDrive -Name HKCC
} # function ConvertRegToScript

function GetWorkspaces {
    [CmdletBinding()]
    param(
        [xml]$XmlData
    )
    $workspaces = $XmlData.respowerfuse.buildingblock.workspaces.workspace
    
    foreach ($workspace in $workspaces) {
        [PsCustomObject]@{
            'Name'    = $workspace.name
            'Guid'    = $workspace.guid
            'Enabled' = $workspace.enabled
        }
    }
} # function GetWorkspaces

function GetPowerzones {
    [CmdletBinding()]
    param(
        [xml]$XmlData
    )
    $workspaces = $XmlData.respowerfuse.buildingblock.powerzones.powerzone
    
    foreach ($workspace in $workspaces) {
        [PsCustomObject]@{
            'Name'    = $workspace.name
            'Guid'    = $workspace.guid
            'Enabled' = $workspace.enabled
        }
    }
} # function GetPowerzones

function GetAccessControls {
    param(
        $Target
    )
    $AccessList = [System.Collections.ArrayList]@()
    $AccessControl = $Target.accesscontrol
    if ($AccessControl.accesstype -eq 'group') { 
        foreach ($group in $AccessControl.grouplist.group) {
            $AccessOptions = @()
            if ($null -ne $group.type) { $AccessOptions += $group.type }
            if ($null -ne $group.InnerText) { $AccessOptions += $group.InnerText }
            $null = $AccessList.Add("$AccessOptions")
        }
        foreach ($group in $AccessControl.notgrouplist.group) {
            $AccessOptions = @()
            if ($null -ne $group.type) { $AccessOptions += $group.type }
            if ($null -ne $group.InnerText) { $AccessOptions += $group.InnerText }
            $null = $AccessList.Add("NOT $AccessOptions")
        }
    }

    #if ($null -ne $AccessControl.access.type) { $null = $AccessList.Add($AccessControl.access.type) }

    if ($null -ne $AccessControl.access) {
        foreach ($subnode in $AccessControl.access) {
            $AccessOptions = @()
            if ($null -ne $subnode.options) { $AccessOptions += $subnode.options }
            if ($null -ne $subnode.type) { $AccessOptions += $subnode.type.ToUpper() }
            if ($null -ne $subnode.object) {
                if ($subnode.type -eq 'powerzone') { 
                    $Name = $Powerzones | Foreach-Object { 
                        if ($_.Guid -eq $subnode.object) { $_.Name } } 
                    }
                else {
                    $Name = $subnode.object
                }
                $AccessOptions += $Name 
            }
            $null = $AccessList.Add("$AccessOptions")
        }
    }
    Write-Output $AccessList
} # function GetAccessControls

function MakeResScript {
    param($script)
        
    $ScriptCommand = ''

    if ($null -ne $script.Workspace) { $WorkspaceText = "    # Workspace: $($script.Workspace)`n" }
    else { $WorkspaceText = '' }

    if ($null -ne $script.name) { $DescriptionText = "    # Name: $($script.name)`n" }
    else { $DescriptionText = '' }

    if ($script.AccessInfo.Count -ne 0) {
        $AccessInfoText = ""
        foreach ($AccessInfo in $script.AccessInfo) {
            $AccessInfoText += "# $AccessInfo`n"
        }
    }
    else {
        $AccessInfoText = ''
    }

    $ScriptCommand += $DescriptionText + $WorkspaceText + $AccessInfoText
    if ($script.Enabled -ne 'yes') {
        $strStart = '# '
        $ScriptCommand += "    # Script was disabled!!!`n"
    }
    else { $strStart = '' }
    if (($script.Type -ne 'ps1')) {
        $strStart = '# '
    }
    $ScriptCommand += "    # Command: $($script.Command) `n"  
    $lines = $script.Text -split "`n"
    foreach ($scriptText in $lines) {
        $ScriptCommand += "    ${strStart}${scriptText} `n"
    } #foreach $scriptText
    Write-Output $ScriptCommand
} # function MakeResScript

function MakeShortcutScript {
    [CmdletBinding()]
    param( [PsCustomObject]$ResObject )
    BEGIN { } #BEGIN
    PROCESS {
        $ScriptTemplate = $script:ScriptTemplate
        $LinkedScriptTemplate = $script:LinkedScriptTemplate

        foreach ($ResEntry in $ResObject) {
            $Target = ''
            $Arguments = ''
            $WorkingDir = ''
            $AppInfo = ''
            if ($ResObject.IsShortcut -eq $True) {
                    
                if (($ResEntry.InStartMenu -eq 'no') -or ($ResEntry.IsEnabled -eq 'no')) {
                    $AppInfo += "    # StartMenu shortcut was disabled!!!`n"
                }

                $Target = [system.environment]::ExpandEnvironmentVariables("$($ResObject.Target)")
                $Arguments = [system.environment]::ExpandEnvironmentVariables("$($ResObject.Arguments)")
                $WorkingDir = [system.environment]::ExpandEnvironmentVariables("$($ResObject.WorkingDir)")
                # insert path to executable
                $StartApp = $script:StartApp
                $StartApp = $StartApp.Replace('<PATHTOAPP>', $Target)
                
                # insert arguments
                if ('' -ne $ResObject.Arguments) {
                    $AppArgs = " -ArgumentList `"$Arguments`""
                }
                else { $AppArgs = '' }
                $StartApp = $StartApp.Replace('<ARGS>', $AppArgs)
    
                # insert working directory
                if ('' -ne $ResObject.WorkingDir) {
                    $AppArgs = " -WorkingDirectory `"$WorkingDir`""
                }
                else { $AppArgs = '' }
                $StartApp = $StartApp.Replace('<WKDIR>', $AppArgs)
                    
                $ScriptBody = $ScriptTemplate
                $ScriptBody = $ScriptBody.Replace('<STARTAPP>', $StartApp)
            } 
            else {
                $ScriptBody = $LinkedScriptTemplate
            }

            # insert info
            if ($ResEntry.AccessInfo.Count -ne 0) {
                foreach ($AccessInfo in $ResEntry.AccessInfo) {
                    $AppInfo += "    # $AccessInfo`n"
                }
            }
                
            # insert linked actipns
            $LinkedScripts = ''
            if ($null -ne $($ResEntry.LinkedApps)) {
                foreach ($LinkedApp in $ResEntry.LinkedApps) { 
                    $LinkedScripts += "    . `"`$PsScriptRoot\__$LinkedApp.ps1`"`n`n"
                } # foreach $LinkedApp
                $LinkedScripts = "$LinkedScripts`n"
            } #if $ResEntry

            # insert environment variables
            $EnvVarCommand = ''
            if ($null -ne $($ResEntry.Variables)) {
                foreach ($var in $ResEntry.Variables) {
                    if ($null -ne $var.description) { $DescriptionText = "    # Description: $($var.description)`n" }
                    else { $DescriptionText = '' }

                    if ($null -ne $var.Workspace) { $WorkspaceText = "    # Workspace: $($var.Workspace)`n" }
                    else { $WorkspaceText = '' }

                    if ($var.AccessInfo.Count -ne 0) {
                        $AccessInfoText = ""
                        foreach ($AccessInfo in $var.AccessInfo) {
                            $AccessInfoText += "    # $AccessInfo`n"
                        }
                    }
                    else {
                        $AccessInfoText = ''
                    }
                        
                    $EnvVarCommand = $EnvVarCommand + $DescriptionText + $WorkspaceText + $AccessInfoText
                        
                    if ($var.Enabled -ne 'yes') {
                        $strStart = '# '
                        $EnvVarCommand += "    # Environment variable was disabled!!! `n"
                    }
                    else { $strStart = '' }
                        
                    $EnvVarCommand += "    ${strStart}`$EnvName  = '$($var.Name)' `n"
                    $EnvVarCommand += "    ${strStart}`$EnvValue = '$($var.Value)' `n"
                    $EnvVarCommand += "    ${strStart}[System.Environment]::SetEnvironmentVariable(`$EnvName, `$EnvValue, `"User`") `n"
                    $EnvVarCommand += "    ${strStart}[System.Environment]::SetEnvironmentVariable(`$EnvName, `$EnvValue, `"Process`")`n"
                    $EnvVarCommand = "$EnvVarCommand`n"
                } # foreach $var
            } # if $ResEntry.Variables

            # insert registry
            $RegistryCommand = ''
            if ($null -ne $($ResEntry.Registry)) {
                foreach ($RegEntry in $ResEntry.Registry) { 
                        
                    if ($null -ne $RegEntry.Workspace) { $WorkspaceText = "    # Workspace: $($RegEntry.Workspace)`n" }
                    else { $WorkspaceText = '' }

                    if ($null -ne $RegEntry.name) { $DescriptionText = "    # Name: $($RegEntry.name)`n" }
                    else { $DescriptionText = '' }

                    if ($RegEntry.AccessInfo.Count -ne 0) {
                        $AccessInfoText = ""
                        foreach ($AccessInfo in $RegEntry.AccessInfo) {
                            $AccessInfoText += "    # $AccessInfo`n"
                        }
                    }
                    else {
                        $AccessInfoText = ''
                    }

                    $RegistryCommand = $RegistryCommand + $DescriptionText + $WorkspaceText + $AccessInfoText
                    if ($RegEntry.Enabled -ne 'yes') {
                        $strStart = '# '
                        $RegistryCommand += "    # Registry was disabled!!! `n"
                    }
                    else { $strStart = '' }
                    foreach ($regString in (ConvertRegToScript -RegData $RegEntry.regText)) {
                        $RegistryCommand += "    ${strStart}${regString} `n"
                    } # foreach $regString
                    $RegistryCommand = "$RegistryCommand`n"
                } # foreach $RegEntry
            } #if $ResEntry
                
            # insert scripts
            $ScriptCommand = ''
            if ($null -ne $($ResEntry.Scripts)) {
                foreach ($script in $ResEntry.Scripts) {
                    $ScriptCommand += MakeResScript -script $script
                    $ScriptCommand = "$ScriptCommand`n"
                } # foreach $ResScript
            } # if $ResEntry.Scripts

            if ($null -ne $($ResEntry.CloseScripts)) {
                foreach ($script in $ResEntry.CloseScripts) {
                    $ScriptCommand += "    # Action on app close!!!`n"
                    $ScriptCommand += MakeResScript -script $script
                    $ScriptCommand = "$ScriptCommand`n"
                } # foreach $ResScript
            } # if $ResEntry.Scripts

        } #foreach ResEntry

        $ScriptBody = $ScriptBody.Replace('<INFO>', $AppInfo)
        $ScriptBody = $ScriptBody.Replace('<LINKEDAPPS>', $LinkedScripts)
        $ScriptBody = $ScriptBody.Replace('<ENVVAR>', $EnvVarCommand)
        $ScriptBody = $ScriptBody.Replace('<REG_ENTRY>', $RegistryCommand)
        $ScriptBody = $ScriptBody.Replace('<SCRIPTS>', $ScriptCommand)

        $ScriptBody
    } #PROCESS
    END { } #END
} #function MakeShortcutScript
}
PROCESS {
foreach ($file in $Path) {
    # enable relative path support
    if ($file -is 'Io.FileInfo') { $file = $file.FullName }
    $file = (resolve-path -Path $file).Path

    Write-Verbose "Parse file $file"
    [xml]$BuildingBlock = Get-Content -Path $file
    
    $Workspaces = GetWorkspaces -XmlData $BuildingBlock
    $PowerZones = GetPowerzones -XmlData $BuildingBlock
    $ResData = ConvertResExport -XmlData $BuildingBlock
        
    if (Test-Path $OutputDir) {} else { $null = New-Item -Path $OutputDir -ItemType Directory }

    foreach ($ResObj in $ResData) {
        $FileNamePrefix = ''
            
        if ($ResObj.IsShortcut -eq $false) {
            $FileNamePrefix = '__'
        }

        if ($null -eq $ResObj.Name) {
            $FileName = (Get-Item $file).BaseName
        }
        else { $FileName = $ResObj.Name }
        $FileName = "$OutputDir\$FileNamePrefix$FileName.ps1"
            
        if ( (Test-Path $FileName) -and !$Force) {
            Write-Verbose "Already exists: $FileName"
            continue
        }
        
        $ShortcutScript = MakeShortcutScript -ResObject $ResObj

        $null = New-Item -Path $FileName -ItemType File -Value $ShortcutScript -Force:$Force
        Write-Verbose "Result file: $FileName"

        if ( $PassThru -and ($null -ne $_.ESNumber) ) { 
            $ResObj | Select-Object ESNumber, Name
        }
    } #foreach $ResObj
} #foreach $file
}
END{}