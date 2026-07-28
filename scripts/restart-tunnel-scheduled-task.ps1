[CmdletBinding()]
param(
  [string]$TaskName = "SharePoint MCP Tunnel",

  [string]$TunnelClientPath = "C:\Apps\TunnelClient\tunnel-client.exe",

  [string]$Profile = "sharepoint-browser-chatgpt"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot "tunnel-task-processes.ps1")

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
if ($task.State -ne "Ready") {
  Stop-ScheduledTask -TaskName $TaskName
  Start-Sleep -Seconds 2
}

$stoppedRoots = Stop-SharePointTunnelProcessTree `
  -TunnelClientPath $TunnelClientPath `
  -Profile $Profile
Start-Sleep -Seconds 2

$remaining = @(
  Get-SharePointTunnelRootProcesses `
    -TunnelClientPath $TunnelClientPath `
    -Profile $Profile
)
if ($remaining.Count -ne 0) {
  throw "Failed to stop $($remaining.Count) SharePoint tunnel process(es)."
}

Start-ScheduledTask -TaskName $TaskName

$running = @()
for ($attempt = 0; $attempt -lt 10; $attempt += 1) {
  Start-Sleep -Seconds 1
  $running = @(
    Get-SharePointTunnelRootProcesses `
      -TunnelClientPath $TunnelClientPath `
      -Profile $Profile
  )
  if ($running.Count -gt 0) {
    break
  }
}

if ($running.Count -ne 1) {
  throw "Expected one SharePoint tunnel process after restart, found $($running.Count)."
}

[pscustomobject]@{
  TaskName = $TaskName
  State = (Get-ScheduledTask -TaskName $TaskName).State
  StoppedTunnelRoots = $stoppedRoots
  RunningTunnelRoots = $running.Count
  ProcessId = $running[0].ProcessId
}
