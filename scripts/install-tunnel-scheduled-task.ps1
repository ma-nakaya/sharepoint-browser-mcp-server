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

Get-ScheduledTask -TaskName $TaskName |
  Select-Object TaskName, State
