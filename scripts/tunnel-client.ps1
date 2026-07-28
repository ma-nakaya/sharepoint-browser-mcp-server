[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidateSet("Init", "Doctor", "Run")]
  [string]$Action,

  [string]$TunnelId,

  [string]$Profile = "sharepoint-browser-chatgpt",

  [string]$TunnelClientPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$serverEntryPoint = Join-Path $repositoryRoot "dist\src\index.js"
$envFile = Join-Path $repositoryRoot ".env"
$profileDirectory = Join-Path $env:APPDATA "tunnel-client"

function Resolve-TunnelClient {
  if ($TunnelClientPath) {
    $resolved = Resolve-Path -LiteralPath $TunnelClientPath -ErrorAction Stop
    return $resolved.Path
  }

  $sharedClient = "C:\Apps\TunnelClient\tunnel-client.exe"
  if (Test-Path -LiteralPath $sharedClient) {
    return $sharedClient
  }

  $localClient = Join-Path $repositoryRoot ".tools\tunnel-client.exe"
  if (Test-Path -LiteralPath $localClient) {
    return $localClient
  }

  $command = Get-Command "tunnel-client" -ErrorAction SilentlyContinue
  if (-not $command) {
    throw "tunnel-client was not found in C:\Apps\TunnelClient, .tools, or PATH. Download the latest release from https://github.com/openai/tunnel-client/releases/latest, or pass -TunnelClientPath."
  }
  return $command.Source
}

function Import-ControlPlaneApiKey {
  if ($env:CONTROL_PLANE_API_KEY -and $env:CONTROL_PLANE_API_KEY.Trim()) {
    return
  }

  if ($env:OPENAI_API_KEY -and $env:OPENAI_API_KEY.Trim()) {
    $env:CONTROL_PLANE_API_KEY = $env:OPENAI_API_KEY
    return
  }

  foreach ($candidateName in @(".env.local", ".env")) {
    $candidatePath = Join-Path $repositoryRoot $candidateName
    if (-not (Test-Path -LiteralPath $candidatePath)) {
      continue
    }

    foreach ($line in Get-Content -LiteralPath $candidatePath) {
      if ($line -match '^\s*OPENAI_API_KEY\s*=\s*(.+?)\s*$') {
        $value = $Matches[1].Trim()
        if (
          ($value.StartsWith('"') -and $value.EndsWith('"')) -or
          ($value.StartsWith("'") -and $value.EndsWith("'"))
        ) {
          $value = $value.Substring(1, $value.Length - 2)
        }
        if ($value) {
          $env:CONTROL_PLANE_API_KEY = $value
          return
        }
      }
    }
  }

  throw "No API key was found. Set CONTROL_PLANE_API_KEY or OPENAI_API_KEY in the process, .env.local, or .env."
}

function Invoke-TunnelClient {
  param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)

  & $script:tunnelClient @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "tunnel-client exited with code $LASTEXITCODE."
  }
}

$tunnelClient = Resolve-TunnelClient
Import-ControlPlaneApiKey

switch ($Action) {
  "Init" {
    if (-not $TunnelId -or $TunnelId -notmatch '^tunnel_[A-Za-z0-9]+$') {
      throw "Init requires -TunnelId with a valid tunnel_... value."
    }
    if (-not (Test-Path -LiteralPath $serverEntryPoint)) {
      throw "The MCP server is not built. Run npm run build first."
    }

    # tunnel-client parses this command with POSIX-style escaping even on Windows.
    # Forward slashes keep absolute Windows paths intact during its preflight.
    $mcpEnvFile = $envFile.Replace("\", "/")
    $mcpEntryPoint = $serverEntryPoint.Replace("\", "/")
    $mcpCommand = 'node --env-file-if-exists="{0}" --enable-source-maps "{1}"' -f `
      $mcpEnvFile, $mcpEntryPoint
    Invoke-TunnelClient `
      "init" `
      "--sample" "sample_mcp_stdio_local" `
      "--profile" $Profile `
      "--profile-dir" $profileDirectory `
      "--tunnel-id" $TunnelId `
      "--health-listen-addr" "127.0.0.1:0" `
      "--mcp-command" $mcpCommand
  }
  "Doctor" {
    Invoke-TunnelClient `
      "doctor" `
      "--profile" $Profile `
      "--profile-dir" $profileDirectory `
      "--explain"
  }
  "Run" {
    Invoke-TunnelClient `
      "run" `
      "--profile" $Profile `
      "--profile-dir" $profileDirectory
  }
}
