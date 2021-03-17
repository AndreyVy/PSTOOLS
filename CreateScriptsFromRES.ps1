function ConvertFrom-ResExport {
    [CmdletBinding()]
param(
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [ValidateScript( { Test-Path $_ })]
    [string[]]$FullName
)
BEGIN {}
PROCESS {
    foreach ($file in $FullName) {
        Write-Verbose "Parse file $file"
        [xml]$BuildingBlock = Get-Content -Path $file
        $applications = $BuildingBlock.respowerfuse.buildingblock.application
        foreach ($app in $applications) {
              
            $EnvVariables = [System.Collections.ArrayList]@()
            foreach ($variable in $app.powerlaunch.variable) {
                #$null = $EnvVariables.Add(@{"$($variable.name)"="$($variable.value)"})
                $null = $EnvVariables.Add("$($variable.name)=$($variable.value)")
            } # foreach $variable
                
            $ResScripts = [System.Collections.ArrayList]@()
            foreach ($script in $app.powerlaunch.exttask) {
                #$null = $ResScripts.Add(@{"$($script.Command)" = "$($script.script)"})
                $null = $ResScripts.Add("$($script.Command) = $($script.script)")
            } # foreach $script

            $ResRegistry = [System.Collections.ArrayList]@()
            foreach ($regEntry in $app.powerlaunch.registry) {
                $regText = [System.Text.Encoding]::ASCII.GetString(
                    [byte[]] -split ($regEntry.registryfile -replace '..', '0x$& ')
                ) # GetString
                $null = $ResRegistry.Add($regText)
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

            $obj = [PsCustomObject][ordered]@{
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
            $obj.PsObject.TypeNames.Insert(0, "ResApplication")
        } # foreach $app
    } #foreach
} #BEGIN
END {}
}

function New-ShortcutScript {
    [CmdletBinding()]
    param()
BEGIN{
$ScriptTemplate = @'
function Start-Application {
    #region // PUT REQUIRED CODE FOR APPLICATION START BELOW THIS LINE // ***************************************
    
    Start-Process -FilePath "<PATHTOAPP>" <# Put application path here #> `
                  <# -ArgumentList "" #>`
                  <# -WorkingDirectory "" #> `
                  -WindowStyle Normal <# POSSIBLE VALUES: Normal Minimized Maximized Hidden #>
    } #endregion // end of required code // *********************************************************************
    #region // PUT ANY REQUIRED ADDITIONAL FUNCTIONS BELOW THIS LINE // *****************************************
    
    
    
    
    #endregion // end of additional functions // ****************************************************************
    #region // DO NOT EDIT TEXT BELOW // ************************************************************************
    Start-Application
    #endregion // NO EDITS // ***********************************************************************************
'@

$EnvVariableTemplate = @'
$EnvName = "<NAME>" # EXAMPLE: $EnvName = "SLBSLS_LICENSE_FILE"
$EnvValue = "<VALUE>" # EXAMPLE: $EnvValue = "port@hostname"
[System.Environment]::SetEnvironmentVariable($EnvName, $EnvValue, "User");
[System.Environment]::SetEnvironmentVariable($EnvName, $EnvValue, "Process");
'@

}
PROCESS{

}
END{}
}