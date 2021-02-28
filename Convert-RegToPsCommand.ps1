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
   [string]$filename
)
$RegistryTypes = @{ "string" = "String";
   "hex"                     = "Binary";
   "dword"                   = "DWord";
   "hex(b)"                  = "QWord";
   "hex(7)"                  = "MultiString";
   "hex(2)"                  = "ExpandString";
   "hex(0)"                  = "Unknown"
}
$RegistryHives = @{ "HKEY_LOCAL_MACHINE\\" = "HKLM:";
   "HKEY_CURRENT_USER\\"                   = "HKCU:";
   "HKEY_CLASSES_ROOT\\"                   = "HKCR:";
   "HKEY_USERS\\"                          = "HKU:";
   "HKEY_CURRENT_CONFIG\\"                 = "HKCC:"
}
$key = ""
$name = ""
$value = ""
$type = ""
$curvalue = ""
$code = @()
$code += "# Proccessed file is ""$filename"""

New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR | Out-Null
New-PSDrive -PSProvider registry -Root HKEY_USERS -Name HKU | Out-Null
New-PSDrive -PSProvider registry -Root HKEY_CURRENT_CONFIG -Name HKCC | Out-Null

Try {
   $getfile = Get-Item $filename
   $file = $getfile.OpenText()
   While (!($file.EndOfStream)) {
      $line = $file.ReadLine()
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
               $code += @"
(Get-Item $parentKey).DeleteSubKeyTree("$subKey")
"@
            }
            else {
               $code += @"
(Get-Item $($parentKey)).CreateSubKey("$($subKey)",`$true) | Out-Null
"@ 
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
               Continue;
            }
            else {
               $value = $Matches[3]
            }
         }
         else {
            # current registry value continues on the next line
            if ($line -match "[\\]$") {
               if (![string]::IsNullOrEmpty($curvalue)) {
                  $curvalue += $line
                  Continue;
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
            $code += @"
(Get-Item $parentKey).OpenSubKey("$subKey",`$true).DeleteValue("$name")
"@
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
            "String" {
               $value = $value -replace '"', '""'
            }
            "Binary" { $value = $value -split "," | ForEach-Object { [System.Convert]::ToInt64($_, 16) } }
            "QWord" {
               $temparray = @()
               $temparray = $value -split ","
               [array]::Reverse($temparray)
               $value = -join $temparray
            }
            "MultiString" {
               $MultiStrings = @()
               $temparray = @()
               $temparray = $value -split ",00,00,00,"
               for ($i = 0; $i -lt ($temparray.Count - 1); $i++) { 
                  $MultiStrings += ([System.Text.Encoding]::Unicode.GetString((($temparray[$i] -split ",") + "00" | ForEach-Object { [System.Convert]::ToInt64($_, 16) }))) -replace '"', '""'
               }
               $value = $MultiStrings
            }
            "ExpandString" {
               if ($value -match "^00,00$") {
                  $value = ""
               }
               else {
                  $value = $value -replace ",00,00$", ""
                  $value = [System.Text.Encoding]::Unicode.GetString((($value -split ",") | ForEach-Object { [System.Convert]::ToInt64($_, 16) }))
                  $value = $value -replace '"', '""'
               }
            }
            "Unknown" {
               $code += @"
# Unknown registry type is not supported [$key]\$name, type "$type"
"@
               Continue;                
            }
                     
         }
         $name = $name -replace '"', '""'
         if (($type -eq "Binary") -or ($type -eq "Unknown")) {
            $value = "@(" + ($value -join ",") + ")"
            $code += @"
(Get-Item $($parentKey)).OpenSubKey("$($subKey)",`$true).SetValue("$($name)",[byte[]] $($value),"$($type)")
"@     
         }
         elseif ($type -eq "MultiString") {
            $value = "@(" + ('"' + ($value -join '","') + '"') + ")"
            $code += @"
(Get-Item $($parentKey)).OpenSubKey("$($subKey)",`$true).SetValue("$($name)",[string[]] $($value),"$($type)")
"@ 
         }    
         elseif (($type -eq "DWord") -or ($type -eq "QWord")) {
            $value = "0x" + $value
            $code += @"
(Get-Item $($parentKey)).OpenSubKey("$($subKey)",`$true).SetValue("$($name)",$($value),"$($type)")
"@     
         }
         else {
            $code += @"
(Get-Item $($parentKey)).OpenSubKey("$($subKey)",`$true).SetValue("$($name)","$($value)","$($type)")
"@     
         }
      }
   }
   Set-Clipboard -Value $code
   $TempFile = New-TemporaryFile
   $code | Out-File -Encoding default $TempFile
   Start-Process notepad $TempFile
   $file.Close()
   Remove-PSDrive -Name HKCR
   Remove-PSDrive -Name HKU
   Remove-PSDrive -Name HKCC
}
catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error') }
