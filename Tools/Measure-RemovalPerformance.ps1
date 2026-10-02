#Requires -Version 7

[CmdletBinding()]
param (
    [string]$BaselineRef = '3712337',
    [ValidateRange(1, 10)]
    [int]$Iterations = 3,
    [ValidateRange(1, 200)]
    [int]$Depth = 80,
    [ValidateRange(1, 100)]
    [int]$Branches = 4,
    [ValidateSet('DeepTree', 'ExcludedTree')]
    [string[]]$Scenario = @('DeepTree'),
    [ValidateRange(1, 1000000)]
    [int]$FileCount = 10000
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
    $excludedRoot = Join-Path $scratch 'excluded-tree'
    if ($Scenario -contains 'ExcludedTree') {
        foreach ($fileIndex in 1..$FileCount) {
            $bucket = Join-Path $excludedRoot "Keep/$([int][math]::Floor(($fileIndex - 1) / 100))"
            $null = [System.IO.Directory]::CreateDirectory($bucket)
            [System.IO.File]::WriteAllText((Join-Path $bucket "$fileIndex.txt"), '')
        }
    }

    foreach ($scenarioName in $Scenario) {
    foreach ($version in $workers.Keys) {
        foreach ($iteration in 0..$Iterations) {
            if ($scenarioName -eq 'DeepTree') {
                $root = Join-Path $scratch "$version-$iteration"
                $null = [System.IO.Directory]::CreateDirectory($root)
                foreach ($branch in 1..$Branches) {
                    $nested = Join-Path $root "branch$branch"
                    foreach ($level in 1..$Depth) { $nested = Join-Path $nested 'd' }
                    $null = [System.IO.Directory]::CreateDirectory($nested)
                }
                $workerParams = @{ Type = 'EmptyFolders'; Path = $root }
                $expectedRemoved = $Branches * ($Depth + 1)
            }
            else {
                $root = $excludedRoot
                [System.IO.File]::WriteAllText((Join-Path $root 'remove.txt'), '')
                $workerParams = @{
                    Type = 'FilesInFolder'; Path = $root; Recurse = $true
                    ExcludeFolder = @(Join-Path $root 'Keep')
                    OlderThanUnit = 'Day'; OlderThanQuantity = 0; OlderThanBasedOn = 'LastWriteTime'
                }
                $expectedRemoved = 1
            }

            $removedCount = 0
            $allocatedBefore = [GC]::GetTotalAllocatedBytes($true)
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            & $workers[$version] @workerParams | ForEach-Object {
                if ($_.Error) { throw "Benchmark removal failed: $($_.FullName): $($_.Error)" }
                if ($_.Action -eq 'Removed') { $removedCount++ }
            }
            $stopwatch.Stop()
            $allocatedBytes = [GC]::GetTotalAllocatedBytes($true) - $allocatedBefore

            if ($removedCount -ne $expectedRemoved) {
                throw "Unexpected removal count: $removedCount"
            }
            if (($scenarioName -eq 'DeepTree') -and [System.IO.Directory]::GetFileSystemEntries($root).Length -ne 0) {
                throw 'Benchmark left unexpected content'
            }
            if (($scenarioName -eq 'ExcludedTree') -and @([System.IO.Directory]::EnumerateFiles($root, '*', 'AllDirectories')).Count -ne $FileCount) {
                throw 'Benchmark changed excluded content'
            }

            if ($iteration -gt 0) {
                [pscustomobject]@{
                    Scenario = $scenarioName
                    Version = $version
                    Iteration = $iteration
                    Removed = $removedCount
                    Milliseconds = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 2)
                    AllocatedMB = [math]::Round($allocatedBytes / 1MB, 2)
                }
            }
            if ($scenarioName -eq 'DeepTree') { [System.IO.Directory]::Delete($root) }
        }
    }
    }
}
finally {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}