function Get-SharePointTunnelRootProcesses {
  param(
    [string]$TunnelClientPath = "C:\Apps\TunnelClient\tunnel-client.exe",

    [string]$Profile = "sharepoint-browser-chatgpt"
  )

  $resolvedClient = [System.IO.Path]::GetFullPath($TunnelClientPath)
  $profilePattern = [regex]::Escape($Profile)
  $argumentPattern = "(?i)(?:^|\s)--profile(?:=|\s+)$profilePattern(?:\s|$)"

  return @(
    Get-CimInstance Win32_Process |
      Where-Object {
        $_.Name -eq "tunnel-client.exe" -and
        $_.ExecutablePath -and
        [System.IO.Path]::GetFullPath($_.ExecutablePath) -eq $resolvedClient -and
        $_.CommandLine -match $argumentPattern
      }
  )
}

function Stop-SharePointTunnelProcessTree {
  param(
    [string]$TunnelClientPath = "C:\Apps\TunnelClient\tunnel-client.exe",

    [string]$Profile = "sharepoint-browser-chatgpt"
  )

  $allProcesses = @(Get-CimInstance Win32_Process)
  $rootProcesses = @(
    Get-SharePointTunnelRootProcesses `
      -TunnelClientPath $TunnelClientPath `
      -Profile $Profile
  )
  if ($rootProcesses.Count -eq 0) {
    return 0
  }

  $depthById = @{}
  foreach ($rootProcess in $rootProcesses) {
    $depthById[[int]$rootProcess.ProcessId] = 0
  }

  do {
    $added = $false
    foreach ($candidate in $allProcesses) {
      $parentId = [int]$candidate.ParentProcessId
      $candidateId = [int]$candidate.ProcessId
      if (
        $depthById.ContainsKey($parentId) -and
        -not $depthById.ContainsKey($candidateId)
      ) {
        $depthById[$candidateId] = [int]$depthById[$parentId] + 1
        $added = $true
      }
    }
  } while ($added)

  $depthById.GetEnumerator() |
    Sort-Object Value -Descending |
    ForEach-Object {
      Stop-Process -Id $_.Key -Force -ErrorAction SilentlyContinue
    }

  return $rootProcesses.Count
}
