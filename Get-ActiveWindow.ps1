<# 
    .Synopsis  
        Gets active window in a user session 
         
    .Description 
        Gets active window in a user session. It displays the process name and window title of the process 
 
    .Example 
        .\Get-ActiveWindow.ps1 
         
        Description 
        ----------- 
        Gets the active window that is currently highlighted. 
 
    .Notes 
        AUTHOR:    Sitaram Pamarthi 
        Website:   http://techibee.com 
#> 
[CmdletBinding()] 
Param( 
) 
Add-Type @" 
  using System; 
  using System.Runtime.InteropServices; 
  public class UserWindows { 
    [DllImport("user32.dll")] 
    public static extern IntPtr GetForegroundWindow(); 
} 
"@ 
try { 
$ActiveHandle = [Windows]::GetForegroundWindow() 
$Process = Get-Process | ? {$_.MainWindowHandle -eq $activeHandle} 
$Process | Select ProcessName, @{Name="AppTitle";Expression= {($_.MainWindowTitle)}} 
} catch { 
    Write-Error "Failed to get active Window details. More Info: $_" 
} 
 