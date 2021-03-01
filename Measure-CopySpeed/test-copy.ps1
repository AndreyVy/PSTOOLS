#.\test-copy.ps1                                               : copy files from current script directory to user %tmp%
#.\test-copy.ps1 -cleanup                                      : remove "test_source_files" from user %tmp%
#.\test-copy.ps1 -target_path D:                               : copy files to D:\TEMP from current script directory
#.\test-copy.ps1 -source_path D:                               : copy files to user %tmp% from D:\Temp
#.\test-copy.ps1 -target_path D:\test_source_files -cleanup    : remove D:\test_source_files
param(
    $source_path,
    $target_path,
    [switch]$cleanup=$false
)
$ScriptCurrentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition

if (!$source_path) { $source_path = Join-Path $ScriptCurrentDirectory "test_source_files" }
if (!$target_path) { $target_path = Join-Path "${env:TMP}" "test_destination_files"}

function copy-job{
    if (!(test-path $source_path)) {
        exit
    }
    Copy-Item -Path $source_path -destination $target_path -Force
}

function cleanup-fileset {
    If(Test-Path $target_path) { Remove-Item -path $target_path -Recurse -Force }
    If(Test-Path $source_path) { Remove-Item -path $source_path -Recurse -Force }
}

if (!$cleanup) { copy-job } else { cleanup-fileset }