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
<#
.SYNOPSIS
    Convert xml object to PsCustomObject
.DESCRIPTION
    Long description
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
    [CmdletBinding()]
    param([xml]$XmlData)
    $applications = $XmlData.respowerfuse.buildingblock.application
    
    $Workspaces = Get-Workspace -XmlData $XmlData

    foreach ($app in $applications) {
        $EnvVariables = [System.Collections.ArrayList]@()
        foreach ($variable in $app.powerlaunch.variable) {
            $null = $EnvVariables.Add(
            [PsCustomObject]@{
                'Name' = "$($variable.name)"
                'Value' = "$($variable.value)"
                "Enabled" = "$($variable.enabled)"
                'Workspace' = $Workspaces | Foreach-Object {if ($_.Guid -in $variable.workspacecontrol.workspace) {$_.Name}}
            })   #   save as hashtable
            #$null = $EnvVariables.Add("$($variable.name)=$($variable.value)")          #   save as string 
        } # foreach $variable
        
        $ResScripts = [System.Collections.ArrayList]@()
        foreach ($script in $app.powerlaunch.exttask) {
            $null = $ResScripts.Add(
            [PsCustomObject]@{
                'Command' = $script.Command
                'Text' = $script.script
                'Type' = $script.scriptext
                'Enabled' = $script.enabled
                'Workspace' = $Workspaces | Foreach-Object {if ($_.Guid -in $script.workspacecontrol.workspace) {$_.Name}}
             })     #   save as object
            #$null = $ResScripts.Add("$($script.Command) = $($script.script)")          #   save as string 
        } # foreach $script

        $ResRegistry = [System.Collections.ArrayList]@()
        foreach ($regEntry in $app.powerlaunch.registry) {
            $null = $ResRegistry.Add(
                [PsCustomObject]@{
                'regText' = [System.Text.Encoding]::ASCII.GetString( [byte[]] -split ($regEntry.registryfile -replace '..', '0x$& ') )
                'Enabled' = $regEntry.enabled
                'Workspace' = $Workspaces | Foreach-Object {if ($_.Guid -in $regEntry.workspacecontrol.workspace) {$_.Name}}
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
        } #hashtable output
    } # foreach $app
} # end function

function New-ShortcutScript {
<#
.SYNOPSIS
    Create shortcut script from pscustom object
.DESCRIPTION
    Long description
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
    [CmdletBinding()]
    param(
        [PsCustomObject]$ResObj,
        [switch]$PassThru
    )
    BEGIN {
        $null = New-Item "$PsScriptRoot\Output" -ItemType Directory -Force

    } #BEGIN
    PROCESS {
        $ScriptTemplate = $script:ScriptTemplate
        $EnvVariableTemplate = $script:EnvVariableTemplate
        $FileName = "$PsScriptRoot\Output\$($ResObj.Name).ps1"
        
        if (Test-Path $FileName) {
            $suffix = "_{0}" -f $((new-guid).guid).split('-') | Select-Object -First 1
            $FileName = "$PsScriptRoot\Output\$($ResObj.Name)$($suffix).ps1"
        }       
        
        Write-Debug "Before ScriptBody"
        $ScriptBody = $ScriptTemplate.Replace('<PATHTOAPP>', $ResObj.Target)
        
        # insert arguments
        if ($ResObj.Arguments) {
            $AppArgs = " -ArgumentList `"$($ResObj.Arguments)`""
        } else { $AppArgs = '' }
        
        $ScriptBody = $ScriptBody.Replace('<ARGS>', $AppArgs)

        # insert working directory
        if ($ResObj.WorkingDir) {
            $AppArgs = " -WorkingDirectory `"$($ResObj.WorkingDir)`""
        }
        else { $AppArgs = '' }
        
        $ScriptBody = $ScriptBody.Replace('<WKDIR>', $AppArgs)

        # insert environment variables
        $varStr = ''
        if ($ResObj.Variables) {
            
            foreach ($var in $ResObj.Variables) {
                
                $varStr = "$varStr `n    # Workspace:`t$($var.Workspace) "
                
                if ($var.Enabled -ne 'yes') {
                    $strStart = '# '
                    $varStr = "$varStr `n    # Variables were disabled!!! "
                } else { $strStart = ''}

                $varStr = "$varStr `n    ${strStart}@{'$($var.Name)' = '$($var.Value)'},"
            } # foreach $var

            $varStr = $varStr.Remove($varStr.Length - 1)
            $varStr = $EnvVariableTemplate.Replace('<ENV_VAR_VALUES>', $varStr)
        }
        $ScriptBody = $ScriptBody.Replace('<ENVVAR>', $varStr)

        # insert registry
        $varStr = ''
        if ($ResObj.Registry) {
            
            foreach ($RegEntry in $ResObj.Registry) {
                
                $varStr = "$varStr `n`n    # Workspace:`t$($RegEntry.Workspace)"

                if ($RegEntry.Enabled -ne 'yes') {
                    $strStart = '# '
                    $varStr = "$varStr `n    # Registry was disabled!!!"
                } else { $strStart = ''}

                foreach ($regString in (ConvertFrom-RegToPS -RegData $RegEntry.regText)) {
                    $varStr = "$varStr `n    ${strStart}${regString}"
                } # foreach $regString

            } # foreach $RegEntry
        }
        $ScriptBody = $ScriptBody.Replace('<REG_ENTRY>', "$varStr`n")

        # insert scripts
        $varStr = ''
        if ($ResObj.Scripts) {
            
            foreach ($script in $ResObj.Scripts) {
                
                $varStr = "$varStr `n`n    # Workspace`t$($script.Workspace)"
                
                if ($script.Enabled -ne 'yes') {
                    $strStart = '# '
                    $varStr = "$varStr `n    # Script was disabled!!!"
                } else { $strStart = ''}

                if (($script.Type -ne 'ps1')) {
                    $strStart = '# '
                }

                $varStr = "$varStr `n    # Command:`t$($script.Command)"
                
                $lines = $script.Text -split "`n"
                foreach ($scriptText in $lines) {
                    $varStr = "$varStr `n    ${strStart}${scriptText}"
                } #foreach $scriptText

            } # foreach $ResScript
        
        } # if $ResObj.Scripts
        
        $ScriptBody = $ScriptBody.Replace('<SCRIPTS>', "$varStr`n")
        $null = New-Item -Path $FileName -ItemType File -Value $ScriptBody

        if ($PassThru) { Write-Output $ResObj }
    } #PROCESS
    END {} #END
} #function New-ShortcutScript

function ConvertFrom-RegToPS {
<#
.Synopsis
   Convert .reg file to PowerShell native code.
.AUTHOR
   Oleksandr Sobakar 24.09.2019
   Oleksandr Sobakar 20.12.2020
.VERSION
   1.0 - initial
   2.0 - new feature: supports removing keys and names like using reg.exe
         major bags fixed: error if name contains some symbols ("[","]","/","=")
         minor bugs fixed       
         new approach in output Powershell code
.DESCRIPTION
   This script converts .reg file into native PowerShell code, allowing to save time on creating, maintaining and updating code.
   Required Parameter -filename - path to .reg file.
   Only native code is produced, no functions.
   If registry key does not exist - script creates it.
   After script is finished, it creates .txt file in %TEMP% with results and opens it in Notepad.
   Also, produced code is copied to Clipboard.
.INSTALLATION
   Run "Install in REG Context Menu.cmd".
   This will copy .ps1 file to %HOMEDRIVE%%HOMEPATH%Cegal\Scripts (usually, H:\Cegal\Scripts) and adds "Convert To PowerShell" to right-click context menu for .reg files.
.EXAMPLE
   Manually:
   Convert_Reg_To_PowerShell.ps1 -filename "xxxx.reg"
.EXAMPLE
   Context-Menu:
   Right-click .reg file and select "Convert To PowerShell".
#>
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
            'Name' = $workspace.name
            'Guid' = $workspace.guid
            'Enabled'= $workspace.enabled
        }
    }
}