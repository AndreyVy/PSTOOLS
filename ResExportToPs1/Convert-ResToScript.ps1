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
    [string[]]$FullName,
    [switch]$OverWrite,
    [switch]$PassThru
)
BEGIN {
    $ScriptTemplate = @'
function Start-Application {
    <ENVVAR>
    <REG_ENTRY>
    <SCRIPTS>
    Start-Process -FilePath "<PATHTOAPP>"<ARGS><WKDIR> -WindowStyle Normal
}
Start-Application
'@

    $EnvVariableTemplate = @'
$EnvVariables = @(<ENV_VAR_VALUES>
)

foreach ($EnvVariable in $EnvVariables) {
    [string]$EnvName = $EnvVariable.Keys
    [string]$EnvValue = $EnvVariable.Values
    [System.Environment]::SetEnvironmentVariable($EnvVariable, $EnvValue, "User");
    [System.Environment]::SetEnvironmentVariable($EnvName, $EnvValue, "Process")
}
'@
    function ConvertFrom-ResExport {
        [CmdletBinding()]
        param([xml]$XmlData)
        $applications = $XmlData.respowerfuse.buildingblock.application
    
        $Workspaces = Get-Workspace -XmlData $XmlData

        foreach ($app in $applications) {
            $EnvVariables = [System.Collections.ArrayList]@()
            foreach ($variable in $app.powerlaunch.variable) {
                $null = $EnvVariables.Add(
                    [PsCustomObject]@{
                        'Name'      = "$($variable.name)"
                        'Value'     = "$($variable.value)"
                        "Enabled"   = "$($variable.enabled)"
                        'Workspace' = $Workspaces | Foreach-Object { if ($_.Guid -in $variable.workspacecontrol.workspace) { $_.Name } }
                    })   #   save as hashtable
                #$null = $EnvVariables.Add("$($variable.name)=$($variable.value)")          #   save as string 
            } # foreach $variable
        
            $ResScripts = [System.Collections.ArrayList]@()
            foreach ($script in $app.powerlaunch.exttask) {
                $null = $ResScripts.Add(
                    [PsCustomObject]@{
                        'Command'   = $script.Command
                        'Text'      = $script.script
                        'Type'      = $script.scriptext
                        'Enabled'   = $script.enabled
                        'Workspace' = $Workspaces | Foreach-Object { if ($_.Guid -in $script.workspacecontrol.workspace) { $_.Name } }
                    })     #   save as object
                #$null = $ResScripts.Add("$($script.Command) = $($script.script)")          #   save as string 
            } # foreach $script

            $ResRegistry = [System.Collections.ArrayList]@()
            foreach ($regEntry in $app.powerlaunch.registry) {
                $null = $ResRegistry.Add(
                    [PsCustomObject]@{
                        'regText'   = [System.Text.Encoding]::ASCII.GetString( [byte[]] -split ($regEntry.registryfile -replace '..', '0x$& ') )
                        'Enabled'   = $regEntry.enabled
                        'Workspace' = $Workspaces | Foreach-Object { if ($_.Guid -in $regEntry.workspacecontrol.workspace) { $_.Name } }
                    }
                )
            } # foreach$regEntry

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

            [PsCustomObject][ordered]@{
                'Name'        = $app.configuration.title
                'Description' = $app.configuration.description
                'Target'      = $app.configuration.commandline
                'Arguments'   = $app.configuration.parameters
                'WorkingDir'  = $app.configuration.workingdir
                'Scripts'     = $ResScripts
                'Variables'   = $EnvVariables
                'Registry'    = $ResRegistry
                'FTA'         = $fta
                'IsShortcut'  = $null -eq $app.required
                'ESNumber'    = $app.accesscontrol.grouplist.group.InnerText
            } #hashtable output
        } # foreach $app
    } # end function

    function New-ShortcutScript {
        [CmdletBinding()]
        param( [PsCustomObject]$ResObject )
        BEGIN { } #BEGIN
        PROCESS {
            # detect res entry responsible for shortcut and create start process command
            $ResEntryPoint = $ResObject | Where-Object { $_.IsShortcut }
            $FileName = "$PsScriptRoot\$($ResEntryPoint.Name).ps1"

            $ScriptTemplate = $script:ScriptTemplate
            $EnvVariableTemplate = $script:EnvVariableTemplate
        
            # insert executable
            $ScriptBody = $ScriptTemplate.Replace('<PATHTOAPP>', $ResEntryPoint.Target)
        
            # insert arguments
            if ($ResObject.Arguments) {
                $AppArgs = " -ArgumentList `"$($ResEntryPoint.Arguments)`""
            }
            else { $AppArgs = '' }
            $ScriptBody = $ScriptBody.Replace('<ARGS>', $AppArgs)

            # insert working directory
            if ($ResObject.WorkingDir) {
                $AppArgs = " -WorkingDirectory `"$($ResEntryPoint.WorkingDir)`""
            }
            else { $AppArgs = '' }
            $ScriptBody = $ScriptBody.Replace('<WKDIR>', $AppArgs)

            $EnvVarCommand = ''
            $RegistryCommand = ''
            $ScriptCommand = ''
            foreach ($ResEntry in $ResObject) {
                # insert environment variables
                if ($ResEntry.Variables) {
                    foreach ($var in $ResEntry.Variables) {
                        $EnvVarCommand = "$EnvVarCommand `n    # Res App:`t`t$($ResEntry.Name)`n    # Workspace:`t$($var.Workspace) "
                        if ($var.Enabled -ne 'yes') {
                            $strStart = '# '
                            $EnvVarCommand = "$EnvVarCommand `n    # Variables were disabled!!! "
                        }
                        else { $strStart = '' }
                        $EnvVarCommand = "$EnvVarCommand `n    ${strStart}@{'$($var.Name)' = '$($var.Value)'},"
                    } # foreach $var
                    $EnvVarCommand = $EnvVarCommand.Remove($EnvVarCommand.Length - 1)
                    $EnvVarCommand = $EnvVariableTemplate.Replace('<ENV_VAR_VALUES>', $EnvVarCommand)
                }

                # insert registry
                if ($ResEntry.Registry) {
                    foreach ($RegEntry in $ResEntry.Registry) { 
                        $RegistryCommand = "$RegistryCommand `n`n    # Res App:`t`t$($ResEntry.Name)`n    # Workspace:`t$($RegEntry.Workspace)"
                        if ($RegEntry.Enabled -ne 'yes') {
                            $strStart = '# '
                            $RegistryCommand = "$RegistryCommand `n    # Registry was disabled!!!"
                        }
                        else { $strStart = '' }
                        foreach ($regString in (ConvertFrom-RegToPS -RegData $RegEntry.regText)) {
                            $RegistryCommand = "$RegistryCommand `n    ${strStart}${regString}"
                        } # foreach $regString
                    } # foreach $RegEntry
                }

                # insert scripts
                if ($ResEntry.Scripts) {
                    foreach ($script in $ResEntry.Scripts) {
                        $ScriptCommand = "$ScriptCommand `n`n    # Res App:`t`t$($ResEntry.Name)`n    # Workspace:`t$($script.Workspace)"
                        if ($script.Enabled -ne 'yes') {
                            $strStart = '# '
                            $ScriptCommand = "$ScriptCommand `n    # Script was disabled!!!"
                        }
                        else { $strStart = '' }
                        if (($script.Type -ne 'ps1')) {
                            $strStart = '# '
                        }
                        $ScriptCommand = "$ScriptCommand `n    # Command:`t$($script.Command)"  
                        $lines = $script.Text -split "`n"
                        foreach ($scriptText in $lines) {
                            $ScriptCommand = "$ScriptCommand `n    ${strStart}${scriptText}"
                        } #foreach $scriptText
                    } # foreach $ResScript
                } # if $ResEntry.Scripts
            } #foreach

            $ScriptBody = $ScriptBody.Replace('<ENVVAR>', $EnvVarCommand)
            $ScriptBody = $ScriptBody.Replace('<REG_ENTRY>', "$RegistryCommand`n")
            $ScriptBody = $ScriptBody.Replace('<SCRIPTS>', "$ScriptCommand`n")
            
            $null = New-Item -Path $FileName -ItemType File -Value $ScriptBody -Force:$script:OverWrite
            Write-Verbose "Result file:`t$FileName"
            if ($script:PassThru) { 
                $ResEntryPoint | Select-Object ESNumber, Name
            }
        } #PROCESS
        END { } #END
    } #function New-ShortcutScript

    function ConvertFrom-RegToPS {
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
                if ($line -match '(^;)|(^Windows Registry Editor)' ) {
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
    }

    function Get-Workspace {
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
    }
}
PROCESS {
    foreach ($file in $FullName) {
        Write-Verbose "Parse file $file"
        [xml]$BuildingBlock = Get-Content -Path $file
        $ResData = ConvertFrom-ResExport -XmlData $BuildingBlock
	    
        New-ShortcutScript -ResObject $ResData      
    } #foreach
} #PROCESS
END {}