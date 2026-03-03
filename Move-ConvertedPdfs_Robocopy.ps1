<#
.SYNOPSIS
Moves PDFs that were successfully created by pub_to_pdf_batch.py, using Robocopy.

.EXAMPLE
.\Move-ConvertedPdfs_Robocopy.ps1 `
  -ManifestCsv "C:\Data\PublisherFiles\conversion_manifest.csv" `
  -SourceRoot "C:\Data\PublisherFiles" `
  -DestinationRoot "D:\Archive\ConvertedPDFs" `
  -WhatIf

  .REAL_MOVE
  .\Move-ConvertedPdfs_Robocopy.ps1 `
  -ManifestCsv "C:\Data\PublisherFiles\conversion_manifest.csv" `
  -SourceRoot "C:\Data\PublisherFiles" `
  -DestinationRoot "D:\Archive\ConvertedPDFs"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$ManifestCsv,

  [Parameter(Mandatory=$true)]
  [string]$SourceRoot,

  [Parameter(Mandatory=$true)]
  [string]$DestinationRoot,

  [int]$ChunkSize = 40,

  [switch]$WhatIf
)

function Get-RelativePath([string]$Root, [string]$FullPath) {
  $rootPath = (Resolve-Path $Root).Path.TrimEnd('\')
  $full = (Resolve-Path $FullPath).Path
  if (-not $full.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Path '$full' is not under SourceRoot '$rootPath'"
  }
  return $full.Substring($rootPath.Length).TrimStart('\')
}

$ManifestCsv = (Resolve-Path $ManifestCsv).Path
$SourceRoot = (Resolve-Path $SourceRoot).Path.TrimEnd('\')
$DestinationRoot = (Resolve-Path $DestinationRoot).Path.TrimEnd('\')

if (-not (Test-Path $DestinationRoot)) {
  if ($WhatIf) {
    Write-Host "[WhatIf] Would create destination root: $DestinationRoot"
  } else {
    New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null
  }
}

$rows = Import-Csv $ManifestCsv

$successRows = $rows | Where-Object { $_.status -eq 'success' -and $_.pdf_path -and (Test-Path $_.pdf_path) }

if (-not $successRows -or $successRows.Count -eq 0) {
  Write-Host "No successful PDF conversions found in manifest (or PDFs missing on disk)."
  exit 0
}

# Group by source directory so Robocopy can be used efficiently
$groups = $successRows | Group-Object {
  Split-Path -Parent $_.pdf_path
}

Write-Host "Found $($successRows.Count) PDFs to move across $($groups.Count) folder group(s)."
Write-Host "SourceRoot:      $SourceRoot"
Write-Host "DestinationRoot: $DestinationRoot"
Write-Host "ChunkSize:       $ChunkSize"
Write-Host ""

foreach ($g in $groups) {
  $sourceDir = $g.Name
  $pdfFiles = $g.Group | ForEach-Object { Split-Path -Leaf $_.pdf_path } | Sort-Object -Unique

  # Destination dir preserves relative path from SourceRoot
  $relDir = Get-RelativePath -Root $SourceRoot -FullPath $sourceDir
  $destDir = Join-Path $DestinationRoot $relDir

  if (-not (Test-Path $destDir)) {
    if ($WhatIf) {
      Write-Host "[WhatIf] Would create: $destDir"
    } else {
      New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }
  }

  # Chunk file arguments to avoid command-line length limits
  for ($i = 0; $i -lt $pdfFiles.Count; $i += $ChunkSize) {
    $chunk = $pdfFiles[$i..([Math]::Min($i + $ChunkSize - 1, $pdfFiles.Count - 1))]

    $args = @(
      "`"$sourceDir`"",
      "`"$destDir`""
    ) + ($chunk | ForEach-Object { "`"$_`"" }) + @(
      "/MOV",          # move files (delete from source on success)
      "/R:1", "/W:1",  # retry behavior
      "/NP",           # no progress
      "/NFL", "/NDL",  # no file list / no dir list
      "/NJH", "/NJS"   # no header / no summary
    )

    $cmd = "robocopy " + ($args -join " ")

    if ($WhatIf) {
      Write-Host "[WhatIf] $cmd"
    } else {
      Write-Host $cmd
      & robocopy @args | Out-Host

      # Robocopy exit codes: 0-7 generally indicate success with various conditions
      $rc = $LASTEXITCODE
      if ($rc -ge 8) {
        Write-Warning "Robocopy reported a failure (exit code $rc) for source '$sourceDir'"
      }
    }
  }
}

Write-Host "Done."
