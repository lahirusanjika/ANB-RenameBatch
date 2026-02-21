<#
.SYNOPSIS
  ANB batch rename tool.
.DESCRIPTION
  Renames files to sequential numbers with optional prefix.
  Supports streaming mode, destination moves, presets, logging, hashing, and restore.
.PARAMETER Path
  Root folder to rename in. Default is current directory.
.PARAMETER Destination
  Optional destination folder to move renamed files into.
.PARAMETER PreserveFolders
  When Destination is set, preserve relative subfolders.
.PARAMETER Prefix
  Optional prefix before the number (for example: "photo_").
.PARAMETER Start
  Starting index (default 1).
.PARAMETER Pad
  Zero-padding width. Use 0 to auto-calc based on count.
.PARAMETER MinPad
  Minimum padding when Pad is 0 (default 4).
.PARAMETER Order
  Sort order: Time, Name, Size, or None (fastest).
.PARAMETER Stream
  Stream mode for huge sets (constant memory). Requires -Order None and no -PerExtension.
.PARAMETER Descending
  Reverse the sort order.
.PARAMETER Recurse
  Include subfolders.
.PARAMETER PerExtension
  Restart numbering per extension.
.PARAMETER Include
  Include patterns (wildcards). Default "*".
.PARAMETER Exclude
  Exclude patterns (wildcards).
.PARAMETER Extensions
  Include only these extensions (dot optional), e.g. jpg, mp4.
.PARAMETER ExcludeExtensions
  Exclude these extensions (dot optional).
.PARAMETER IncludeHidden
  Include hidden/system files.
.PARAMETER OnConflict
  Target exists behavior: Error, Skip, or Overwrite.
.PARAMETER Force
  Skip large-batch confirmation.
.PARAMETER ConfirmThreshold
  Number of files that triggers confirmation (default 10000).
.PARAMETER ProgressEvery
  Update progress every N files (0 disables).
.PARAMETER DryRun
  Preview rename plan without changing files.
.PARAMETER ShowAll
  Show all preview lines (otherwise first 50).
.PARAMETER LogPath
  Write a CSV log of actual changes (OriginalFullName, FinalFullName, HashAlgorithm, Hash).
.PARAMETER NoLog
  Disable auto log creation.
.PARAMETER RestoreFrom
  CSV log to restore (undo). Must contain OriginalFullName and FinalFullName.
.PARAMETER HashAlgorithm
  Hash algorithm for logging and restore verification (default SHA256).
.PARAMETER NoHash
  Disable hashing and hash verification.
.PARAMETER SkipHashMismatch
  Skip files with hash mismatches during restore (otherwise error).
.PARAMETER Preset
  Load a JSON preset file.
.PARAMETER SavePreset
  Save current settings to a JSON preset and exit.
.PARAMETER Interactive
  Prompt for options interactively.
.EXAMPLE
  ./tools/ANB-RenameBatch.ps1 -Path . -Order Time
.EXAMPLE
  ./tools/ANB-RenameBatch.ps1 -Path . -Destination D:\Sorted -PreserveFolders
.EXAMPLE
  ./tools/ANB-RenameBatch.ps1 -Order None -Stream -Pad 6
.EXAMPLE
  ./tools/ANB-RenameBatch.ps1 -RestoreFrom .\ANB_RenameLog_20260221_180000.csv
.EXAMPLE
  ./tools/ANB-RenameBatch.ps1 -SavePreset .\preset.json
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
  [Parameter(Position = 0)]
  [string]$Path = ".",
  [string]$Destination,
  [switch]$PreserveFolders,
  [string]$Prefix = "",
  [int]$Start = 1,
  [int]$Pad = 0,
  [int]$MinPad = 4,
  [ValidateSet("Time", "Name", "Size", "None")]
  [string]$Order = "Time",
  [switch]$Stream,
  [switch]$Descending,
  [switch]$Recurse,
  [switch]$PerExtension,
  [string[]]$Include = @("*"),
  [string[]]$Exclude = @(),
  [string[]]$Extensions = @(),
  [string[]]$ExcludeExtensions = @(),
  [switch]$IncludeHidden,
  [ValidateSet("Error", "Skip", "Overwrite")]
  [string]$OnConflict = "Error",
  [switch]$Force,
  [int]$ConfirmThreshold = 10000,
  [int]$ProgressEvery = 1000,
  [switch]$DryRun,
  [switch]$ShowAll,
  [string]$LogPath,
  [switch]$NoLog,
  [string]$RestoreFrom,
  [ValidateSet("SHA256", "SHA1", "MD5", "SHA384", "SHA512")]
  [string]$HashAlgorithm = "SHA256",
  [switch]$NoHash,
  [switch]$SkipHashMismatch,
  [string]$Preset,
  [string]$SavePreset,
  [switch]$Interactive
)

function Test-AnyLike {
  param(
    [string]$Value,
    [string[]]$Patterns
  )
  foreach ($p in $Patterns) {
    if ($Value -like $p) { return $true }
  }
  return $false
}

function Normalize-Extensions {
  param([string[]]$List)
  $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($item in $List) {
    if ($null -eq $item) { continue }
    $e = $item.Trim()
    if ($e.Length -eq 0 -or $e -eq ".") { continue }
    if ($e[0] -ne ".") { $e = "." + $e }
    $null = $set.Add($e.ToLowerInvariant())
  }
  return $set
}

function Escape-CsvValue {
  param([string]$Value)
  if ($null -eq $Value) { return '""' }
  '"' + ($Value -replace '"', '""') + '"'
}

function Build-CsvLine {
  param([string[]]$Values)
  ($Values | ForEach-Object { Escape-CsvValue $_ }) -join ','
}
function Resolve-ANBPath {
  param([string]$InputPath)
  if ([string]::IsNullOrWhiteSpace($InputPath)) { return $null }
  try {
    return (Resolve-Path -LiteralPath $InputPath -ErrorAction Stop).Path
  } catch {
    return [System.IO.Path]::GetFullPath($InputPath)
  }
}

function Get-RootDir {
  param([string]$InputPath)
  $resolved = Resolve-ANBPath $InputPath
  if (-not $resolved) { return (Get-Location).Path }
  if (Test-Path -LiteralPath $resolved -PathType Leaf) {
    return (Split-Path -Parent $resolved)
  }
  return $resolved
}

function Get-RelativePathSafe {
  param([string]$FromPath, [string]$ToPath)
  $mi = [System.IO.Path].GetMethod("GetRelativePath", [type[]]@([string], [string]))
  if ($mi) {
    return [System.IO.Path]::GetRelativePath($FromPath, $ToPath)
  }
  $from = ([System.IO.Path]::GetFullPath($FromPath)).TrimEnd("\") + "\"
  $to = ([System.IO.Path]::GetFullPath($ToPath)).TrimEnd("\") + "\"
  $fromUri = [System.Uri]::new($from)
  $toUri = [System.Uri]::new($to)
  $rel = $fromUri.MakeRelativeUri($toUri).ToString()
  $rel = [System.Uri]::UnescapeDataString($rel).Replace('/', '\').TrimEnd('\')
  if ([string]::IsNullOrWhiteSpace($rel)) { return "." }
  return $rel
}

function Get-FinalDirectory {
  param([System.IO.FileInfo]$File)
  if (-not $script:DestinationResolved) { return $File.DirectoryName }
  if (-not $script:PreserveFolders) { return $script:DestinationResolved }
  $rel = Get-RelativePathSafe $script:RootDir $File.DirectoryName
  if ($rel -eq "." -or [string]::IsNullOrWhiteSpace($rel)) { return $script:DestinationResolved }
  return (Join-Path -Path $script:DestinationResolved -ChildPath $rel)
}

function Get-ANBItems {
  foreach ($item in Get-ChildItem -Path $Path -File -Recurse:$Recurse -Force:$IncludeHidden) {
    if ($script:ExcludeDestination -and $item.FullName.StartsWith($script:DestinationNorm, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
    if ($script:IncludeAll -eq $false -and -not (Test-AnyLike $item.Name $Include)) { continue }
    if ($script:HasExclude -and (Test-AnyLike $item.Name $Exclude)) { continue }
    $ext = $item.Extension.ToLowerInvariant()
    if ($script:HasIncludeExt -and -not $script:IncludeExtSet.Contains($ext)) { continue }
    if ($script:HasExcludeExt -and $script:ExcludeExtSet.Contains($ext)) { continue }
    $item
  }
}

function Get-ANBHash {
  param([string]$PathToHash, [string]$Algorithm)
  if ($script:NoHash) { return $null }
  try {
    return (Get-FileHash -LiteralPath $PathToHash -Algorithm $Algorithm).Hash
  } catch {
    throw "Failed to hash file: $PathToHash"
  }
}

function Write-ANBLogHeader {
  param(
    [System.IO.StreamWriter]$Writer,
    [switch]$IncludeHash
  )
  $cols = @("OriginalFullName", "FinalFullName")
  if ($IncludeHash) { $cols += @("HashAlgorithm", "Hash") }
  $Writer.WriteLine((Build-CsvLine $cols))
}

function Write-ANBLogRow {
  param(
    [System.IO.StreamWriter]$Writer,
    [string]$Original,
    [string]$Final,
    [string]$Algorithm,
    [string]$Hash,
    [switch]$IncludeHash
  )
  if ($IncludeHash) {
    $Writer.WriteLine((Build-CsvLine @($Original, $Final, $Algorithm, $Hash)))
  } else {
    $Writer.WriteLine((Build-CsvLine @($Original, $Final)))
  }
}
function Resolve-ANBLogPath {
  if ($script:LogPath) { return $script:LogPath }
  if ($script:NoLog) { return $null }
  if ($script:UsePreview) { return $null }
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $logDir = $null
  try {
    $resolved = Resolve-Path -LiteralPath $script:Path -ErrorAction Stop | Select-Object -First 1
    $logDir = $resolved.Path
  } catch {
    $logDir = $null
  }
  if ($logDir -and (Test-Path -LiteralPath $logDir -PathType Leaf)) {
    $logDir = Split-Path -Parent $logDir
  }
  if (-not $logDir) { $logDir = (Get-Location).Path }
  return (Join-Path -Path $logDir -ChildPath ("ANB_RenameLog_{0}.csv" -f $ts))
}

function Confirm-ANBBatch {
  param([int]$Count, [string]$Action)
  if ($script:UsePreview) { return $true }
  if ($script:Force) { return $true }
  if ($Count -ge $script:ConfirmThreshold) {
    return $PSCmdlet.ShouldContinue("$Action will process $Count files. Continue?", "Large batch confirmation")
  }
  return $PSCmdlet.ShouldProcess("$Action $Count files", "ANB Rename Batch")
}

function Write-ANBProgress {
  param([string]$Activity, [int]$Index, [int]$Total)
  if ($script:ProgressEvery -le 0) { return }
  if (($Index % $script:ProgressEvery -eq 0) -or ($Index -eq $Total)) {
    $percent = if ($Total -gt 0) { [math]::Floor(($Index / $Total) * 100) } else { 0 }
    Write-Progress -Activity $Activity -Status "$Index / $Total" -PercentComplete $percent
  }
}

function Invoke-ANBRenameRollback {
  param([object[]]$Map)
  Write-Warning "Rollback started for rename operation."
  foreach ($m in $Map) {
    try {
      if (Test-Path -LiteralPath $m.FinalFullName) {
        if ((Test-Path -LiteralPath $m.OriginalFullName) -and $script:OnConflict -ne "Overwrite") {
          Write-Warning "Rollback skip (target exists): $($m.OriginalFullName)"
        } else {
          Move-Item -LiteralPath $m.FinalFullName -Destination $m.OriginalFullName -Force
        }
      } elseif (Test-Path -LiteralPath $m.TempFullName) {
        if ((Test-Path -LiteralPath $m.OriginalFullName) -and $script:OnConflict -ne "Overwrite") {
          Write-Warning "Rollback skip (target exists): $($m.OriginalFullName)"
        } else {
          Rename-Item -LiteralPath $m.TempFullName -NewName (Split-Path -Leaf $m.OriginalFullName)
        }
      }
    } catch {
      Write-Warning "Rollback error: $($_.Exception.Message)"
    }
  }
}

function Invoke-ANBRestoreRollback {
  param(
    [object[]]$Map,
    [object[]]$MovedToOriginal = @()
  )
  Write-Warning "Rollback started for restore operation."
  $movedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($moved in $MovedToOriginal) {
    if ($moved -and $moved.OriginalFullName) { $null = $movedSet.Add($moved.OriginalFullName) }
  }
  foreach ($m in $Map) {
    try {
      if (Test-Path -LiteralPath $m.TempFullName) {
        $dir = Split-Path -Parent $m.FinalFullName
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        Move-Item -LiteralPath $m.TempFullName -Destination $m.FinalFullName -Force
        continue
      }
      if ($movedSet.Contains($m.OriginalFullName) -and (Test-Path -LiteralPath $m.OriginalFullName)) {
        $dir = Split-Path -Parent $m.FinalFullName
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        Move-Item -LiteralPath $m.OriginalFullName -Destination $m.FinalFullName -Force
      }
    } catch {
      Write-Warning "Rollback error: $($_.Exception.Message)"
    }
  }
}
function Apply-ANBPreset {
  param([hashtable]$Data, [hashtable]$Bound)
  foreach ($key in $Data.Keys) {
    if ($Bound.ContainsKey($key)) { continue }
    if (Get-Variable -Name $key -Scope Script -ErrorAction SilentlyContinue) {
      $val = $Data[$key]
      if ($null -ne $val -and $val.GetType().Name -eq "PSCustomObject") {
        $val = $val | ConvertTo-Json -Depth 5 | ConvertFrom-Json
      }
      Set-Variable -Name $key -Scope Script -Value $val
    }
  }
}

function Save-ANBPreset {
  param([string]$PresetPath)
  $data = [ordered]@{
    Path = $Path
    Destination = $Destination
    PreserveFolders = [bool]$PreserveFolders
    Prefix = $Prefix
    Start = $Start
    Pad = $Pad
    MinPad = $MinPad
    Order = $Order
    Stream = [bool]$Stream
    Descending = [bool]$Descending
    Recurse = [bool]$Recurse
    PerExtension = [bool]$PerExtension
    Include = $Include
    Exclude = $Exclude
    Extensions = $Extensions
    ExcludeExtensions = $ExcludeExtensions
    IncludeHidden = [bool]$IncludeHidden
    OnConflict = $OnConflict
    Force = [bool]$Force
    ConfirmThreshold = $ConfirmThreshold
    ProgressEvery = $ProgressEvery
    DryRun = [bool]$DryRun
    ShowAll = [bool]$ShowAll
    LogPath = $LogPath
    NoLog = [bool]$NoLog
    HashAlgorithm = $HashAlgorithm
    NoHash = [bool]$NoHash
    SkipHashMismatch = [bool]$SkipHashMismatch
  }
  $json = $data | ConvertTo-Json -Depth 6
  $dir = Split-Path -Parent $PresetPath
  if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
  }
  Set-Content -LiteralPath $PresetPath -Value $json -Encoding UTF8
}

function Read-ANBString {
  param([string]$Label, [string]$Default)
  $suffix = if ([string]::IsNullOrWhiteSpace($Default)) { "" } else { " [$Default]" }
  $input = Read-Host "$Label$suffix"
  if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
  return $input
}

function Read-ANBInt {
  param([string]$Label, [int]$Default)
  $input = Read-Host "$Label [$Default]"
  if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
  $parsed = 0
  if ([int]::TryParse($input, [ref]$parsed)) { return $parsed }
  return $Default
}

function Read-ANBYesNo {
  param([string]$Label, [bool]$Default)
  $def = if ($Default) { "Y" } else { "N" }
  $input = Read-Host "$Label (Y/N) [$def]"
  if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
  switch -Regex ($input.Trim()) {
    "^(y|yes)$" { return $true }
    "^(n|no)$" { return $false }
    default { return $Default }
  }
}

function Read-ANBChoice {
  param([string]$Label, [string[]]$Choices, [string]$Default)
  $opts = $Choices -join "/"
  $input = Read-Host "$Label ($opts) [$Default]"
  if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
  if ($Choices -contains $input) { return $input }
  return $Default
}

function Read-ANBList {
  param([string]$Label, [string[]]$Default)
  $def = if ($Default) { ($Default -join ",") } else { "" }
  $input = Read-Host "$Label (comma, '-' for none) [$def]"
  if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
  if ($input.Trim() -eq "-") { return @() }
  return ($input -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
}

$ErrorActionPreference = "Stop"

$bound = @{}
foreach ($k in $PSBoundParameters.Keys) { $bound[$k] = $true }

if ($Preset) {
  if (-not (Test-Path -LiteralPath $Preset)) { throw "Preset not found: $Preset" }
  $presetObj = Get-Content -LiteralPath $Preset -Raw | ConvertFrom-Json
  $presetData = @{}
  foreach ($prop in $presetObj.PSObject.Properties) {
    $presetData[$prop.Name] = $prop.Value
  }
  Apply-ANBPreset -Data $presetData -Bound $bound
}

if ($Interactive) {
  $Path = Read-ANBString "Path" $Path
  $Destination = Read-ANBString "Destination (blank for none)" $Destination
  if ([string]::IsNullOrWhiteSpace($Destination)) { $Destination = $null }
  if ($Destination) {
    $PreserveFolders = Read-ANBYesNo "Preserve folders" ([bool]$PreserveFolders)
  }
  $Order = Read-ANBChoice "Order" @("Time", "Name", "Size", "None") $Order
  $Stream = Read-ANBYesNo "Stream mode" ([bool]$Stream)
  if ($Stream) { $Order = "None" }
  $Recurse = Read-ANBYesNo "Recurse" ([bool]$Recurse)
  $PerExtension = Read-ANBYesNo "Per extension sequences" ([bool]$PerExtension)
  $Prefix = Read-ANBString "Prefix" $Prefix
  $Start = Read-ANBInt "Start index" $Start
  $Pad = Read-ANBInt "Pad width (0 auto)" $Pad
  $MinPad = Read-ANBInt "Min pad width" $MinPad
  $Extensions = Read-ANBList "Extensions include" $Extensions
  $ExcludeExtensions = Read-ANBList "Extensions exclude" $ExcludeExtensions
  $OnConflict = Read-ANBChoice "OnConflict" @("Error", "Skip", "Overwrite") $OnConflict
  $ProgressEvery = Read-ANBInt "Progress every N (0 disables)" $ProgressEvery
  $DryRun = Read-ANBYesNo "Dry run" ([bool]$DryRun)
  $hashEnabled = Read-ANBYesNo "Enable hashing" (-not $NoHash)
  $NoHash = -not $hashEnabled
  if (-not $NoHash) {
    $HashAlgorithm = Read-ANBChoice "Hash algorithm" @("SHA256", "SHA1", "SHA384", "SHA512", "MD5") $HashAlgorithm
  }
}

if ($SavePreset) {
  Save-ANBPreset -PresetPath $SavePreset
  Write-Output "Preset saved to $SavePreset"
  return
}

if ($Start -lt 0) { throw "Start must be >= 0." }
if ($Pad -lt 0) { throw "Pad must be >= 0." }
if ($MinPad -lt 1) { throw "MinPad must be >= 1." }
if ($ProgressEvery -lt 0) { throw "ProgressEvery must be >= 0." }
if ($ConfirmThreshold -lt 0) { throw "ConfirmThreshold must be >= 0." }

if ($Stream) {
  if ($Order -ne "None") { throw "Stream mode requires -Order None." }
  if ($PerExtension) { throw "Stream mode does not support -PerExtension." }
}

$RootDir = Get-RootDir $Path
$DestinationResolved = Resolve-ANBPath $Destination
if ($DestinationResolved) {
  if (-not (Test-Path -LiteralPath $DestinationResolved)) {
    New-Item -ItemType Directory -Force -Path $DestinationResolved | Out-Null
  }
}

$IncludeAll = ($Include.Count -eq 1 -and $Include[0] -eq "*")
$HasExclude = ($Exclude -and $Exclude.Count -gt 0)
$IncludeExtSet = Normalize-Extensions $Extensions
$ExcludeExtSet = Normalize-Extensions $ExcludeExtensions
$HasIncludeExt = ($IncludeExtSet.Count -gt 0)
$HasExcludeExt = ($ExcludeExtSet.Count -gt 0)

$DestinationNorm = $null
$ExcludeDestination = $false
if ($DestinationResolved) {
  $rootNorm = ([System.IO.Path]::GetFullPath($RootDir)).TrimEnd('\') + '\'
  $destNorm = ([System.IO.Path]::GetFullPath($DestinationResolved)).TrimEnd('\') + '\'
  if ($destNorm.StartsWith($rootNorm, [System.StringComparison]::OrdinalIgnoreCase) -and $destNorm -ne $rootNorm) {
    $ExcludeDestination = $true
    $DestinationNorm = $destNorm
  }
}

$usePreview = $DryRun -or $WhatIfPreference
$script:UsePreview = $usePreview
if ($RestoreFrom) {
  if (-not (Test-Path -LiteralPath $RestoreFrom)) {
    throw "Restore log not found: $RestoreFrom"
  }

  $rows = Import-Csv -LiteralPath $RestoreFrom
  if (-not $rows -or $rows.Count -eq 0) {
    throw "Restore log is empty."
  }

  $props = $rows[0].PSObject.Properties.Name
  if (-not ($props -contains "OriginalFullName") -or -not ($props -contains "FinalFullName")) {
    throw "Restore log must contain columns: OriginalFullName, FinalFullName."
  }
  $hasHash = ($props -contains "Hash")
  $hasAlgo = ($props -contains "HashAlgorithm")

  $map = New-Object System.Collections.Generic.List[object]
  foreach ($r in $rows) {
    if ([string]::IsNullOrWhiteSpace($r.OriginalFullName) -or [string]::IsNullOrWhiteSpace($r.FinalFullName)) { continue }
    $map.Add([pscustomobject]@{
      OriginalFullName = $r.OriginalFullName
      FinalFullName = $r.FinalFullName
      Hash = if ($hasHash) { $r.Hash } else { $null }
      HashAlgorithm = if ($hasAlgo) { $r.HashAlgorithm } else { $null }
    })
  }

  if ($map.Count -eq 0) { throw "Restore log has no valid rows." }

  $origSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($m in $map) {
    if (-not $origSet.Add($m.OriginalFullName)) {
      throw "Collision: multiple entries target the same original path: $($m.OriginalFullName)"
    }
  }

  if (-not (Confirm-ANBBatch -Count $map.Count -Action "Restore")) { return }

  if ($usePreview) {
    Write-Output "Preview restore: $($map.Count) files"
    $list = if ($ShowAll) { $map } else { $map | Select-Object -First 50 }
    foreach ($m in $list) {
      "{0} -> {1}" -f $m.FinalFullName, $m.OriginalFullName
    }
    if (-not $ShowAll -and $map.Count -gt 50) {
      Write-Output "... (use -ShowAll to see full list)"
    }
    return
  }

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $guid = [guid]::NewGuid().ToString("N")
  $tempPrefix = "__ANB_RESTORE_TMP_$guid"
  $tempIndex = 0
  $restoreMap = New-Object System.Collections.Generic.List[object]

  $i = 0
  try {
    foreach ($m in $map) {
      if (-not (Test-Path -LiteralPath $m.FinalFullName)) {
        if ($OnConflict -eq "Skip") {
          Write-Warning "Missing source, skipping: $($m.FinalFullName)"
          continue
        }
        throw "Missing source: $($m.FinalFullName)"
      }

      if (-not $NoHash -and $hasHash -and -not [string]::IsNullOrWhiteSpace($m.Hash)) {
        $algo = if ([string]::IsNullOrWhiteSpace($m.HashAlgorithm)) { $HashAlgorithm } else { $m.HashAlgorithm }
        $actual = Get-ANBHash -PathToHash $m.FinalFullName -Algorithm $algo
        if ($actual -ne $m.Hash) {
          if ($SkipHashMismatch) {
            Write-Warning "Hash mismatch, skipping: $($m.FinalFullName)"
            continue
          }
          throw "Hash mismatch: $($m.FinalFullName)"
        }
      }

      $tempIndex++
      $ext = [System.IO.Path]::GetExtension($m.FinalFullName)
      $tempName = "{0}_{1}{2}" -f $tempPrefix, $tempIndex.ToString("D6"), $ext
      $tempFull = Join-Path -Path (Split-Path -Parent $m.FinalFullName) -ChildPath $tempName
      if (Test-Path -LiteralPath $tempFull) {
        throw "Temp file already exists (likely from a previous run): $tempFull"
      }
      Rename-Item -LiteralPath $m.FinalFullName -NewName $tempName
      $restoreMap.Add([pscustomobject]@{
        OriginalFullName = $m.OriginalFullName
        TempFullName = $tempFull
        FinalFullName = $m.FinalFullName
      })
      $i++
      Write-ANBProgress -Activity "Restore (pass 1/2)" -Index $i -Total $map.Count
    }
  } catch {
    Invoke-ANBRestoreRollback -Map $restoreMap
    throw
  }
  if ($ProgressEvery -gt 0) { Write-Progress -Activity "Restore (pass 1/2)" -Completed }

  $logWriter = $null
  if ($LogPath) {
    $logWriter = New-Object System.IO.StreamWriter($LogPath, $false, [System.Text.Encoding]::UTF8)
    Write-ANBLogHeader -Writer $logWriter -IncludeHash:$false
  }

  $i = 0
  $restoredToOriginal = New-Object System.Collections.Generic.List[object]
  try {
    foreach ($m in $restoreMap) {
      $i++
      $skip = $false
      $origDir = Split-Path -Parent $m.OriginalFullName
      if ($origDir -and -not (Test-Path -LiteralPath $origDir)) {
        New-Item -ItemType Directory -Force -Path $origDir | Out-Null
      }
      if (Test-Path -LiteralPath $m.OriginalFullName) {
        switch ($OnConflict) {
          "Skip" {
            Move-Item -LiteralPath $m.TempFullName -Destination $m.FinalFullName
            Write-Warning "Target exists, skipped: $($m.OriginalFullName)"
            Write-ANBProgress -Activity "Restore (pass 2/2)" -Index $i -Total $restoreMap.Count
            $skip = $true
            break
          }
          "Overwrite" {
            Move-Item -LiteralPath $m.TempFullName -Destination $m.OriginalFullName -Force
            $restoredToOriginal.Add($m)
          }
          Default {
            throw "Target already exists: $($m.OriginalFullName)"
          }
        }
        if ($skip) { continue }
      } else {
        Move-Item -LiteralPath $m.TempFullName -Destination $m.OriginalFullName
        $restoredToOriginal.Add($m)
      }
      if ($logWriter) {
        Write-ANBLogRow -Writer $logWriter -Original $m.OriginalFullName -Final $m.FinalFullName -IncludeHash:$false
      }
      Write-ANBProgress -Activity "Restore (pass 2/2)" -Index $i -Total $restoreMap.Count
    }
  } catch {
    if ($logWriter) { $logWriter.Dispose() }
    Invoke-ANBRestoreRollback -Map $restoreMap -MovedToOriginal $restoredToOriginal
    throw
  }
  if ($logWriter) { $logWriter.Dispose() }
  if ($ProgressEvery -gt 0) { Write-Progress -Activity "Restore (pass 2/2)" -Completed }

  $sw.Stop()
  $duration = $sw.Elapsed.ToString()
  $msg = "Restored $($restoreMap.Count) files in $duration."
  if ($LogPath) { $msg += " Log: $LogPath" }
  Write-Output $msg
  return
}
# Rename mode
$logPathResolved = Resolve-ANBLogPath

if ($Stream) {
  $count = 0
  foreach ($f in Get-ANBItems) { $count++ }
  if ($count -eq 0) { throw "No files matched." }
  if (-not (Confirm-ANBBatch -Count $count -Action "Rename")) { return }

  $width = if ($Pad -gt 0) { $Pad } else { [math]::Max($MinPad, ($Start + $count - 1).ToString().Length) }

  if ($usePreview) {
    if ($logPathResolved) {
      $logWriter = New-Object System.IO.StreamWriter($logPathResolved, $false, [System.Text.Encoding]::UTF8)
      try {
        Write-ANBLogHeader -Writer $logWriter -IncludeHash:$false
        $shown = 0
        foreach ($f in Get-ANBItems) {
          $number = $Start + $shown
          $base = $number.ToString("D$width")
          if ($Prefix) { $base = "$Prefix$base" }
          $destDir = Get-FinalDirectory -File $f
          $finalFull = Join-Path -Path $destDir -ChildPath ("$base$($f.Extension)")
          if ($ShowAll -or $shown -lt 50) {
            "{0} -> {1}" -f $f.FullName, $finalFull
          }
          Write-ANBLogRow -Writer $logWriter -Original $f.FullName -Final $finalFull -IncludeHash:$false
          $shown++
        }
      } finally {
        $logWriter.Dispose()
      }
    } else {
      $shown = 0
      Write-Output "Preview: $count files"
      foreach ($f in Get-ANBItems) {
        $number = $Start + $shown
        $base = $number.ToString("D$width")
        if ($Prefix) { $base = "$Prefix$base" }
        $destDir = Get-FinalDirectory -File $f
        $finalFull = Join-Path -Path $destDir -ChildPath ("$base$($f.Extension)")
        if ($ShowAll -or $shown -lt 50) {
          "{0} -> {1}" -f $f.FullName, $finalFull
        } else {
          break
        }
        $shown++
      }
      if (-not $ShowAll -and $count -gt 50) {
        Write-Output "... (use -ShowAll or -LogPath to see full list)"
      }
    }
    return
  }

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $guid = [guid]::NewGuid().ToString("N")
  $tempPrefix = "__ANB_TMP_$guid"
  $mapPath = Join-Path -Path $env:TEMP -ChildPath ("ANB_RenameMap_{0}.csv" -f $guid)

  $mapWriter = $null
  try {
    $mapWriter = New-Object System.IO.StreamWriter($mapPath, $false, [System.Text.Encoding]::UTF8)
    $mapWriter.WriteLine((Build-CsvLine @("OriginalFullName", "TempFullName", "FinalFullName")))
    $index = 0
    foreach ($f in Get-ANBItems) {
      $number = $Start + $index
      $base = $number.ToString("D$width")
      if ($Prefix) { $base = "$Prefix$base" }
      $destDir = Get-FinalDirectory -File $f
      $finalFull = Join-Path -Path $destDir -ChildPath ("$base$($f.Extension)")

      $tempName = "{0}_{1}{2}" -f $tempPrefix, ($index + 1).ToString("D6"), $f.Extension
      $tempFull = Join-Path -Path $f.DirectoryName -ChildPath $tempName
      if (Test-Path -LiteralPath $tempFull) {
        throw "Temp file already exists (likely from a previous run): $tempFull"
      }
      Rename-Item -LiteralPath $f.FullName -NewName $tempName
      $mapWriter.WriteLine((Build-CsvLine @($f.FullName, $tempFull, $finalFull)))
      $index++
      Write-ANBProgress -Activity "Rename (pass 1/2)" -Index $index -Total $count
    }
  } catch {
    if ($mapWriter) { $mapWriter.Dispose() }
    if (Test-Path -LiteralPath $mapPath) {
      $rollbackMap = Import-Csv -LiteralPath $mapPath
      Invoke-ANBRenameRollback -Map $rollbackMap
    }
    throw
  } finally {
    if ($mapWriter) { $mapWriter.Dispose() }
  }
  if ($ProgressEvery -gt 0) { Write-Progress -Activity "Rename (pass 1/2)" -Completed }

  $logWriter = $null
  if ($logPathResolved) {
    $logWriter = New-Object System.IO.StreamWriter($logPathResolved, $false, [System.Text.Encoding]::UTF8)
    Write-ANBLogHeader -Writer $logWriter -IncludeHash:(-not $NoHash)
  }

  $index = 0
  try {
    foreach ($m in Import-Csv -LiteralPath $mapPath) {
      $index++
      $skip = $false
      $hash = $null
      if (-not $NoHash) { $hash = Get-ANBHash -PathToHash $m.TempFullName -Algorithm $HashAlgorithm }
      if (Test-Path -LiteralPath $m.FinalFullName) {
        switch ($OnConflict) {
          "Skip" {
            Move-Item -LiteralPath $m.TempFullName -Destination $m.OriginalFullName
            Write-Warning "Target exists, skipped: $($m.FinalFullName)"
            Write-ANBProgress -Activity "Rename (pass 2/2)" -Index $index -Total $count
            $skip = $true
            break
          }
          "Overwrite" {
            $dir = Split-Path -Parent $m.FinalFullName
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
            Move-Item -LiteralPath $m.TempFullName -Destination $m.FinalFullName -Force
          }
          Default {
            throw "Target already exists and is not part of this rename set: $($m.FinalFullName)"
          }
        }
        if ($skip) { continue }
      } else {
        $dir = Split-Path -Parent $m.FinalFullName
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        if ($OnConflict -eq "Overwrite") {
          Move-Item -LiteralPath $m.TempFullName -Destination $m.FinalFullName -Force
        } else {
          Move-Item -LiteralPath $m.TempFullName -Destination $m.FinalFullName
        }
      }
      if ($logWriter) {
        Write-ANBLogRow -Writer $logWriter -Original $m.OriginalFullName -Final $m.FinalFullName -Algorithm $HashAlgorithm -Hash $hash -IncludeHash:(-not $NoHash)
      }
      Write-ANBProgress -Activity "Rename (pass 2/2)" -Index $index -Total $count
    }
  } catch {
    if ($logWriter) { $logWriter.Dispose() }
    $rollbackMap = Import-Csv -LiteralPath $mapPath
    Invoke-ANBRenameRollback -Map $rollbackMap
    throw
  }
  if ($logWriter) { $logWriter.Dispose() }
  if ($ProgressEvery -gt 0) { Write-Progress -Activity "Rename (pass 2/2)" -Completed }

  Remove-Item -LiteralPath $mapPath -ErrorAction SilentlyContinue
  $sw.Stop()
  $duration = $sw.Elapsed.ToString()
  $msg = "Renamed $count files in $duration."
  if ($logPathResolved) { $msg += " Log: $logPathResolved" }
  Write-Output $msg
  return
}
$items = Get-ANBItems
if (-not $items -or $items.Count -eq 0) {
  throw "No files matched."
}

if ($Order -ne "None") {
  switch ($Order) {
    "Time" { $props = @("LastWriteTime", "Name") }
    "Name" { $props = @("Name") }
    "Size" { $props = @("Length", "Name") }
  }
  $items = $items | Sort-Object -Property $props -Descending:$Descending
}

$groups = if ($PerExtension) {
  $items | Group-Object { $_.Extension.ToLowerInvariant() }
} else {
  @([pscustomobject]@{ Name = "all"; Group = $items })
}

$guid = [guid]::NewGuid().ToString("N")
$tempPrefix = "__ANB_TMP_$guid"
$globalIndex = 0
$map = New-Object System.Collections.Generic.List[object]

foreach ($g in $groups) {
  $groupFiles = @($g.Group)
  $groupCount = $groupFiles.Count
  $width = if ($Pad -gt 0) { $Pad } else { [math]::Max($MinPad, ($Start + $groupCount - 1).ToString().Length) }

  for ($i = 0; $i -lt $groupCount; $i++) {
    $f = $groupFiles[$i]
    $number = $Start + $i
    $base = $number.ToString("D$width")
    if ($Prefix) { $base = "$Prefix$base" }
    $destDir = Get-FinalDirectory -File $f
    $finalName = "$base$($f.Extension)"
    $finalFull = Join-Path -Path $destDir -ChildPath $finalName

    $globalIndex++
    $tempName = "{0}_{1}{2}" -f $tempPrefix, $globalIndex.ToString("D6"), $f.Extension
    $tempFull = Join-Path -Path $f.DirectoryName -ChildPath $tempName

    $map.Add([pscustomobject]@{
      OriginalFullName = $f.FullName
      TempFullName = $tempFull
      FinalFullName = $finalFull
    })
  }
}

$comparer = [System.StringComparer]::OrdinalIgnoreCase
$originalSet = [System.Collections.Generic.HashSet[string]]::new($comparer)
$finalSet = [System.Collections.Generic.HashSet[string]]::new($comparer)
foreach ($m in $map) {
  $null = $originalSet.Add($m.OriginalFullName)
  if (-not $finalSet.Add($m.FinalFullName)) {
    throw "Collision: multiple files map to the same target: $($m.FinalFullName)"
  }
  if (Test-Path -LiteralPath $m.TempFullName) {
    throw "Temp file already exists (likely from a previous run): $($m.TempFullName)"
  }
}

$filtered = New-Object System.Collections.Generic.List[object]
foreach ($m in $map) {
  if (Test-Path -LiteralPath $m.FinalFullName) {
    if (-not $originalSet.Contains($m.FinalFullName)) {
      switch ($OnConflict) {
        "Skip" {
          Write-Warning "Target exists, skipped: $($m.FinalFullName)"
          continue
        }
        "Overwrite" {
          $filtered.Add($m)
          continue
        }
        Default {
          throw "Target already exists and is not part of this rename set: $($m.FinalFullName)"
        }
      }
    }
  }
  $filtered.Add($m)
}
$map = $filtered

if ($map.Count -eq 0) { throw "No files to rename after conflict handling." }
if (-not (Confirm-ANBBatch -Count $map.Count -Action "Rename")) { return }

if ($usePreview) {
  Write-Output "Preview: $($map.Count) files"
  $list = if ($ShowAll) { $map } else { $map | Select-Object -First 50 }
  foreach ($m in $list) {
    "{0} -> {1}" -f $m.OriginalFullName, $m.FinalFullName
  }
  if (-not $ShowAll -and $map.Count -gt 50) {
    Write-Output "... (use -ShowAll or -LogPath to see full list)"
  }
  if ($logPathResolved) {
    $logWriter = New-Object System.IO.StreamWriter($logPathResolved, $false, [System.Text.Encoding]::UTF8)
    Write-ANBLogHeader -Writer $logWriter -IncludeHash:$false
    foreach ($m in $map) {
      Write-ANBLogRow -Writer $logWriter -Original $m.OriginalFullName -Final $m.FinalFullName -IncludeHash:$false
    }
    $logWriter.Dispose()
  }
  return
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$pass1List = New-Object System.Collections.Generic.List[object]

$i = 0
try {
  foreach ($m in $map) {
    Rename-Item -LiteralPath $m.OriginalFullName -NewName (Split-Path -Leaf $m.TempFullName)
    $pass1List.Add($m)
    $i++
    Write-ANBProgress -Activity "Rename (pass 1/2)" -Index $i -Total $map.Count
  }
} catch {
  Invoke-ANBRenameRollback -Map $pass1List
  throw
}
if ($ProgressEvery -gt 0) { Write-Progress -Activity "Rename (pass 1/2)" -Completed }

$logWriter = $null
if ($logPathResolved) {
  $logWriter = New-Object System.IO.StreamWriter($logPathResolved, $false, [System.Text.Encoding]::UTF8)
  Write-ANBLogHeader -Writer $logWriter -IncludeHash:(-not $NoHash)
}

$i = 0
try {
  foreach ($m in $map) {
    $hash = $null
    if (-not $NoHash) { $hash = Get-ANBHash -PathToHash $m.TempFullName -Algorithm $HashAlgorithm }
    $dir = Split-Path -Parent $m.FinalFullName
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    if ($OnConflict -eq "Overwrite") {
      Move-Item -LiteralPath $m.TempFullName -Destination $m.FinalFullName -Force
    } else {
      Move-Item -LiteralPath $m.TempFullName -Destination $m.FinalFullName
    }
    if ($logWriter) {
      Write-ANBLogRow -Writer $logWriter -Original $m.OriginalFullName -Final $m.FinalFullName -Algorithm $HashAlgorithm -Hash $hash -IncludeHash:(-not $NoHash)
    }
    $i++
    Write-ANBProgress -Activity "Rename (pass 2/2)" -Index $i -Total $map.Count
  }
} catch {
  if ($logWriter) { $logWriter.Dispose() }
  Invoke-ANBRenameRollback -Map $map
  throw
}
if ($logWriter) { $logWriter.Dispose() }
if ($ProgressEvery -gt 0) { Write-Progress -Activity "Rename (pass 2/2)" -Completed }

$sw.Stop()
$duration = $sw.Elapsed.ToString()
$msg = "Renamed $($map.Count) files in $duration."
if ($logPathResolved) { $msg += " Log: $logPathResolved" }
Write-Output $msg
