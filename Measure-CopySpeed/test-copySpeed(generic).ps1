param(
    $amount = 50
    , $size = 1mb
    , $name = "dummy"

)
function random-datafile {
    param(
        $file_path,
        $file_size
    )
    $data = new-object byte[] $file_size
    (new-object Random).NextBytes($data)
    [IO.File]::WriteAllBytes($file_path, $data) 
}
function emty-datafile {
    param (
        $file_path,
        $file_size
    )
    $file = [io.file]::Create($file_path)
    $file.SetLength($file_size)
    $file.Close() | Out-Null
}

function test-rnd {
    
    $source_path_rnd = Join-Path $ScriptCurrentDirectory "rnd_files"
    $target_path_rnd = Join-Path "${env:TMP}" "results_rnd"

    new-item -path $source_path_rnd -itemtype Directory -Force | out-null
    
    $rnd_start_create = Get-Date

    for ($i = 0; $i -lt $amount; $i++) {
        $file_path = join-path $source_path_rnd "$([guid]::NewGuid()).rnd"
        random-datafile -file_path $file_path -file_size $size
    }

    $rnd_stop_create = Get-Date
    
    Move-Item -Path $source_path_rnd -destination $target_path_rnd -Force

    $rnd_stop_move = Get-Date

    #Write-host "Create RND Files (sec): " ($rnd_stop_create - $rnd_start_create).TotalSeconds
    #Write-host "Move RND File (sec): " ($rnd_stop_move - $rnd_stop_create).TotalSeconds

    [math]::Round(($rnd_stop_create - $rnd_start_create).TotalSeconds, 3)
    [math]::Round(($rnd_stop_move - $rnd_stop_create).TotalSeconds, 3)
}

function test-empty {
   
    $source_path_empty = Join-Path $ScriptCurrentDirectory "empty_files"
    $target_path_empty = Join-Path "${env:TMP}" "results_empty"  

    new-item -path $source_path_empty -itemtype Directory -Force | out-null

    $empty_start_create = Get-Date

    for ($i = 0; $i -lt $amount; $i++) {
        $file_path = join-path $source_path_empty "$([guid]::NewGuid()).emp"
        emty-datafile -file_path $file_path -file_size $size
    }

    $empty_stop_create = Get-Date
    
    Move-Item -Path $source_path_empty -destination $target_path_empty -Force
    $empty_stop_move = Get-Date
        
    #Write-host "Create EMPTY Files: " ($empty_stop_create - $empty_start_create).TotalSeconds
    #Write-host "Move EMPTY Files: " ($empty_stop_move - $empty_stop_create).TotalSeconds

    [math]::Round(($empty_stop_create - $empty_start_create).TotalSeconds, 3)
    [math]::Round(($empty_stop_move - $empty_stop_create).TotalSeconds, 3)
}

$ScriptCurrentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition

$rnd_create = @()
$rnd_move = @()

$empty_create = @()
$empty_move = @()

for ($j = 0; $j -lt 3; $j++) {
    $result = test-rnd
    $rnd_create += $result[0]
    $rnd_move += $result[1]
    
    $result = test-empty
    $empty_create += $result[0]
    $empty_move += $result[1]
}

Write-Host "`nCreate RND Files" 
$rnd_create | Format-Table
Write-Host "`nMove RND Files"
$rnd_move | Format-Table

Write-Host "`nCreate EMPTY Files"
$empty_create | Format-Table
Write-Host "`nMove EMPTY Files"
$empty_move | Format-Table