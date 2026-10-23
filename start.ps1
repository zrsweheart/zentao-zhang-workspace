param(
  [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'
$port = 8765
$healthUrl = "http://127.0.0.1:$port/api/health"
$workspaceUrl = "http://127.0.0.1:$port/"
$appPath = Join-Path $PSScriptRoot 'app.py'
$pidPath = Join-Path $PSScriptRoot '.server.pid'
$stdoutPath = Join-Path $PSScriptRoot 'server.log'
$stderrPath = Join-Path $PSScriptRoot 'server-error.log'

$python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) {
  $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
  if (-not $pythonCommand) {
    throw 'Python was not found. Create .venv and install requirements.txt, or install Python 3.10+.'
  }
  $python = $pythonCommand.Source
}

function Test-WorkspaceServer {
  try {
    $result = Invoke-RestMethod -Uri $healthUrl -TimeoutSec 2
    return $result.ok -eq $true
  } catch {
    return $false
  }
}

function Get-ListeningProcessId {
  try {
    $connection = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction Stop | Select-Object -First 1
    if ($connection) { return [int]$connection.OwningProcess }
  } catch {
    $line = netstat -ano | Select-String "127\.0\.0\.1:$port\s+.*LISTENING\s+(\d+)" | Select-Object -First 1
    if ($line -and $line.Matches.Count) { return [int]$line.Matches[0].Groups[1].Value }
  }
  return 0
}

if (Test-WorkspaceServer) {
  $oldPid = Get-ListeningProcessId
  if ($oldPid -gt 0 -and $oldPid -ne $PID) {
    Write-Host "Stopping previous workspace process ($oldPid)..."
    Stop-Process -Id $oldPid -Force
    for ($i = 0; $i -lt 20; $i++) {
      if (-not (Test-WorkspaceServer)) { break }
      Start-Sleep -Milliseconds 200
    }
  }
}

$env:PYTHONUTF8 = '1'
$processInfo = New-Object System.Diagnostics.ProcessStartInfo
$processInfo.FileName = $python
$processInfo.Arguments = '"' + $appPath + '"'
$processInfo.WorkingDirectory = $PSScriptRoot
$processInfo.UseShellExecute = $true
$processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
$process = [System.Diagnostics.Process]::Start($processInfo)
Set-Content -LiteralPath $pidPath -Value $process.Id -Encoding ascii

for ($i = 0; $i -lt 30; $i++) {
  if (Test-WorkspaceServer) {
    Write-Host "Workspace started: $workspaceUrl"
    if (-not $NoBrowser) { Start-Process $workspaceUrl }
    exit 0
  }
  if ($process.HasExited) {
    throw 'Workspace exited during startup. Run app.py in a PowerShell window to see the Python error.'
  }
  Start-Sleep -Milliseconds 300
}

throw "Workspace did not become ready. Check $stderrPath"
