function Get-DnsNameByIp {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [string[]]$IpAddress
    )
    BEGIN {}
    PROCESS {

        ForEach ($Addr in $Address) {
            $props = @{'IpAddress'=$addr}
            Try {
                $result = [System.Net.Dns]::GetHostByAddress($addr)
                $props.Add('ComputerName',$result.HostName)
            } Catch {
                $props.Add('ComputerName',$null)
            }
            New-Object -TypeName PSObject -Property $props
        } #foreach

    } #PROCESS
    END {}

} #function