[CmdletBinding()]
param(
  [string]$TaskName = "SharePoint MCP Tunnel",

  [string]$AppsDirectory = "C:\Apps\TunnelClient"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$launcherSource = Join-Path $PSScriptRoot "run-sharepoint-mcp-tunnel.vbs"
$launcherDestination = Join-Path $AppsDirectory "run-sharepoint-mcp-tunnel.vbs"
$tunnelClient = Join-Path $AppsDirectory "tunnel-client.exe"
$profileDirectory = Join-Path $env:APPDATA "tunnel-client"
$profilePath = Join-Path $profileDirectory "sharepoint-browser-chatgpt.yaml"
$controlPlaneApiKey = [Environment]::GetEnvironmentVariable(
  "CONTROL_PLANE_API_KEY",
  "User"
)

. (Join-Path $PSScriptRoot "tunnel-task-processes.ps1")

if (-not (Test-Path -LiteralPath $launcherSource)) {
  throw "VBS launcher was not found: $launcherSource"
}
if (-not (Test-Path -LiteralPath $tunnelClient)) {
  throw "tunnel-client was not found: $tunnelClient"
}
if (-not (Test-Path -LiteralPath $profilePath)) {
  throw "Tunnel profile was not found: $profilePath"
}
if (-not $controlPlaneApiKey -or -not $controlPlaneApiKey.Trim()) {
  throw "CONTROL_PLANE_API_KEY is not set in the current user's environment."
}

$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask -and $existingTask.State -ne "Ready") {
  Stop-ScheduledTask -TaskName $TaskName
  Start-Sleep -Seconds 2
}
Stop-SharePointTunnelProcessTree -TunnelClientPath $tunnelClient | Out-Null

Copy-Item -LiteralPath $launcherSource -Destination $launcherDestination -Force

$userId = "$env:USERDOMAIN\$env:USERNAME"
$action = New-ScheduledTaskAction `
  -Execute (Join-Path $env:SystemRoot "System32\wscript.exe") `
  -Argument "`"$launcherDestination`"" `
  -WorkingDirectory $AppsDirectory
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $userId
$principal = New-ScheduledTaskPrincipal `
  -UserId $userId `
  -LogonType Interactive `
  -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries `
  -DontStopIfGoingOnBatteries `
  -ExecutionTimeLimit ([TimeSpan]::Zero) `
  -Hidden `
  -MultipleInstances IgnoreNew `
  -RestartCount 3 `
  -RestartInterval (New-TimeSpan -Minutes 1) `
  -StartWhenAvailable

Register-ScheduledTask `
  -TaskName $TaskName `
  -Action $action `
  -Trigger $trigger `
  -Principal $principal `
  -Settings $settings `
  -Description "Runs the SharePoint MCP Secure Tunnel without a visible console window." `
  -Force | Out-Null

Start-ScheduledTask -TaskName $TaskName

$running = @()
for ($attempt = 0; $attempt -lt 10; $attempt += 1) {
  Start-Sleep -Seconds 1
  $running = @(Get-SharePointTunnelRootProcesses -TunnelClientPath $tunnelClient)
  if ($running.Count -gt 0) {
    break
  }
}
if ($running.Count -ne 1) {
  throw "Expected one SharePoint tunnel process after installation, found $($running.Count)."
}

[pscustomobject]@{
  TaskName = $TaskName
  State = (Get-ScheduledTask -TaskName $TaskName).State
  RunningTunnelRoots = $running.Count
  ProcessId = $running[0].ProcessId
}
