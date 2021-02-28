function Get-Netstat {
    [CmdletBinding()]
    param(
        [switch]$infitiny = $false
    )
    BEGIN{ }
    PROCESS{
        do {
            netstat -a -n -o | Select-Object -Skip 4 | 
                        ConvertFrom-String -PropertyNames Blank, Protocol,LocalAddress, ForeignAddress, State, PID |
                        Select-Object Protocol,
                                    @{name='LocalAddress'; expression={($_.LocalAddress).Substring(0,($_.LocalAddress).LastIndexOf(':') )}},
                                    @{name='LocalPort'; expression={(($_.LocalAddress) -split ":")[-1]}},
                                    @{name='ForeignAddress'; expression={($_.ForeignAddress).Substring(0,($_.ForeignAddress).LastIndexOf(':') )}},
                                    @{'Name'='ForeignPort'; expression={(($_.ForeignAddress) -split ":")[-1]}},
                                    State,
                                    @{'Name'='ProcessName'; expression={If ($_.PID) {(Get-Process -Id $_.PID).Name} else {$Null}}}
        } while ($infitiny)
        
    }
    END{ }
}