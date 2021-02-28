function Get-Assoc {
    [CmdLetBinding(DefaultParameterSetName="Program")]
    Param(
        [parameter(ParameterSetName="Program")]
        [string]$ProgramPath="*"
    )
    $AssocResults = &"cmd" -Args "/c" "assoc" |
                    ConvertFrom-String -Delimiter "=" -PropertyNames Extension, ProgId 
    $ftypeResults = &"cmd" -Args "/c" "ftype" |
                    ConvertFrom-String -Delimiter "=" -PropertyNames ProgId, Command |
                    Where-Object Command -Like $ProgramPath
  
    foreach ($ftypeResult in $ftypeResults) {
        $HasExtensions = $AssocResults | Where-Object ProgId -eq $ftypeResult.ProgId
        if ($null -eq $HasExtensions) {continue}
        foreach ($HasExtension in $HasExtensions) {

                $Description = (&"cmd" -Args "/c" "assoc $($ftypeResult.ProgId) 2> nul" |
                ConvertFrom-String -Delimiter "=" -PropertyNames ProgId, Description).Description
            if ($null -eq $Description) { $Description = $ftypeResult.ProgId }
            
            $Command = $ftypeResult.Command
            
                $props = @{
                'Extension' = $HasExtension.Extension
                'Description' = $Description
                'Command' = $Command
            }
            New-Object -TypeName PSObject -Property $props
        }
    }
}