$port = 43741
$dataRoot = Join-Path (Get-Location) "output\packaged-smoke-data"
$env:INVENTORY_NO_BROWSER = "1"
$env:INVENTORY_PORT = [string]$port
$env:INVENTORY_DATA_ROOT = $dataRoot

$process = Start-Process -FilePath ".\release\win-unpacked\resources\backend\inventory_backend.exe" -PassThru -WindowStyle Hidden
try {
    $state = $null
    for ($attempt = 0; $attempt -lt 30 -and -not $state; $attempt++) {
        try {
            $state = Invoke-RestMethod -Uri ("http://127.0.0.1:" + $port + "/api/state") -TimeoutSec 1
        } catch {
            Start-Sleep -Milliseconds 250
        }
    }
    if (-not $state.ok) {
        throw "Packaged backend did not return a healthy API response."
    }
    Write-Output ("packaged backend smoke test: OK (pid " + $process.Id + ")")
} finally {
    if (-not $process.HasExited) {
        Stop-Process -Id $process.Id -Force
    }
}
