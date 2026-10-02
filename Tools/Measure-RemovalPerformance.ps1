#Requires -Version 7

[CmdletBinding()]
param (
    [string]$BaselineRef = '3712337',
    [ValidateRange(1, 10)]
    [int]$Iterations = 3,
    [ValidateRange(1, 200)]
    [int]$Depth = 80,
    [ValidateRange(1, 100)]
    [int]$Branches = 4
)

$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot
$baselineText = git -C $repository show "${BaselineRef}:Remove items.ps1"
if ($LASTEXITCODE -ne 0) { throw "Cannot read baseline '$BaselineRef'" }
$workers = [ordered]@{
    Baseline = [scriptblock]::Create($baselineText -join [Environment]::NewLine)
    Current = [scriptblock]::Create((Get-Content -LiteralPath "$repository/Remove items.ps1" -Raw))
}
$scratch = Join-Path ([System.IO.Path]::GetTempPath()) "RemovalBenchmark_$([guid]::NewGuid())"
$null = [System.IO.Directory]::CreateDirectory($scratch)

try {
    foreach ($version in $workers.Keys) {
        foreach ($iteration in 0..$Iterations) {
            $root = Join-Path $scratch "$version-$iteration"
            $null = [System.IO.Directory]::CreateDirectory($root)
            foreach ($branch in 1..$Branches) {
                $nested = Join-Path $root "branch$branch"
                foreach ($level in 1..$Depth) { $nested = Join-Path $nested 'd' }
                $null = [System.IO.Directory]::CreateDirectory($nested)
            }

            $removedCount = 0
            $allocatedBefore = [GC]::GetTotalAllocatedBytes($true)
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            & $workers[$version] -Type EmptyFolders -Path $root | ForEach-Object {
                if ($_.Error) { throw "Benchmark removal failed: $($_.FullName): $($_.Error)" }
                if ($_.Action -eq 'Removed') { $removedCount++ }
            }
            $stopwatch.Stop()
            $allocatedBytes = [GC]::GetTotalAllocatedBytes($true) - $allocatedBefore

            if ($removedCount -ne ($Branches * ($Depth + 1))) {
                throw "Unexpected removal count: $removedCount"
            }
            if ([System.IO.Directory]::GetFileSystemEntries($root).Length -ne 0) {
                throw 'Benchmark left unexpected content'
            }

            if ($iteration -gt 0) {
                [pscustomobject]@{
                    Scenario = 'DeepTree'
                    Version = $version
                    Iteration = $iteration
                    Removed = $removedCount
                    Milliseconds = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 2)
                    AllocatedMB = [math]::Round($allocatedBytes / 1MB, 2)
                }
            }
            [System.IO.Directory]::Delete($root)
        }
    }
}
finally {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}