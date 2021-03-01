# .\new-files.ps1               : create 100 1mb-files in current script directory\test_source_files
# .\new-files.ps1 -cleanup      : remove current script directory\test_source_files
# .\new-files.ps1 -amount 20    : create 20 1mb-files in current script directory\test_source_files
# .\new-files.ps1 -size 1gb     : create 100 1gb-files in current script directory\test_source_files
param(
    $amount = 100,
    $size = 1mb,
    [switch]$cleanup=$false
)

$ScriptCurrentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
$source_path = Join-Path $ScriptCurrentDirectory "test_source_files"

function new-emptyfile {
    param (
        $file_path,
        $file_size
    )
    $file = [io.file]::Create($file_path)
    $file.SetLength($file_size)
    $file.Close() | Out-Null
}

function new-fileset {
    new-item -path $source_path -itemtype Directory -Force | out-null
    for ($i = 0; $i -lt $amount; $i++) {
        $file_path = join-path $source_path "$([guid]::NewGuid()).emp"
        new-emptyfile -file_path $file_path -file_size $size
    }
}

function cleanup-fileset {
    If(Test-Path $source_path) { Remove-Item -path $source_path -Recurse -Force }
}

if (!$cleanup) { new-fileset } else { cleanup-fileset }