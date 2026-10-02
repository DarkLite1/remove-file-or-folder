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
    [ValidateSet('DeepTree', 'ExcludedTree', 'FileFiltering', 'WideDirectory', 'FileDeletion')]
    [string[]]$Scenario = @('DeepTree', 'ExcludedTree', 'FileFiltering', 'WideDirectory', 'FileDeletion'),
    [ValidateRange(1, 1000000)]
    [int]$FileCount = 10000,
    [ValidateRange(0, 1000000)]
    [int]$ExcludeFileCount = 250
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
    foreach ($scenarioName in $Scenario) {
        $scenarioRoot = Join-Path $scratch $scenarioName
        $filePaths = [System.Collections.Generic.List[string]]::new()
        $excludedFilePaths = [System.Collections.Generic.List[string]]::new()
        if ($scenarioName -ne 'DeepTree') {
            foreach ($fileIndex in 1..$FileCount) {
                $bucketName = if ($scenarioName -eq 'WideDirectory') {
                    'Files'
                }
                else { "Keep/$([int][math]::Floor(($fileIndex - 1) / 100))" }
                $bucket = Join-Path $scenarioRoot $bucketName
                $null = [System.IO.Directory]::CreateDirectory($bucket)
                $filePath = Join-Path $bucket "$fileIndex.txt"
                [System.IO.File]::WriteAllText($filePath, '')
                $filePaths.Add($filePath)
                if (($scenarioName -eq 'FileFiltering') -and ($fileIndex -le $ExcludeFileCount)) {
                    [System.IO.File]::SetLastWriteTime($filePath, [datetime]::Now.AddYears(-2))
                    $excludedFilePaths.Add($filePath)
                }
            }
        }

        foreach ($version in $workers.Keys) {
            foreach ($iteration in 0..$Iterations) {
                $root = $scenarioRoot
                $expectedRemoved = 0
                $expectedRemaining = $FileCount
                $workerParams = @{
                    Type              = 'FilesInFolder'
                    Path              = $root
                    Recurse           = $true
                    OlderThanUnit     = 'Day'
                    OlderThanQuantity = 0
                    OlderThanBasedOn  = 'LastWriteTime'
                }
                switch ($scenarioName) {
                    'DeepTree' {
                        $root = Join-Path $scenarioRoot "$version-$iteration"
                        foreach ($branch in 1..$Branches) {
                            $nested = Join-Path $root "branch$branch"
                            foreach ($level in 1..$Depth) { $nested = Join-Path $nested 'd' }
                            $null = [System.IO.Directory]::CreateDirectory($nested)
                        }
                        $workerParams = @{ Type = 'EmptyFolders'; Path = $root }
                        $expectedRemoved = $Branches * ($Depth + 1)
                        $expectedRemaining = 0
                    }
                    'ExcludedTree' {
                        [System.IO.File]::WriteAllText((Join-Path $root 'remove.txt'), '')
                        $workerParams.ExcludeFolder = @(Join-Path $root 'Keep')
                        $expectedRemoved = 1
                    }
                    'FileFiltering' {
                        $workerParams.OlderThanQuantity = 30
                        $workerParams.ExcludeFile = $excludedFilePaths.ToArray()
                    }
                    'WideDirectory' {
                        $workerParams = @{ Type = 'EmptyFolders'; Path = $root }
                    }
                    'FileDeletion' {
                        foreach ($filePath in $filePaths) { [System.IO.File]::WriteAllText($filePath, '') }
                        $expectedRemoved = $FileCount
                        $expectedRemaining = 0
                    }
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

                if ($removedCount -ne $expectedRemoved) { throw "Unexpected removal count: $removedCount" }
                if (@([System.IO.Directory]::EnumerateFiles($root, '*', 'AllDirectories')).Count -ne $expectedRemaining) {
                    throw 'Benchmark changed retained content or left files behind'
                }
                if (($scenarioName -eq 'DeepTree') -and [System.IO.Directory]::GetFileSystemEntries($root).Length -ne 0) {
                    throw 'Benchmark left unexpected directories'
                }

                if ($iteration -gt 0) {
                    [pscustomobject]@{
                        Scenario     = $scenarioName
                        Version      = $version
                        Iteration    = $iteration
                        Removed      = $removedCount
                        Milliseconds = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 2)
                        AllocatedMB  = [math]::Round($allocatedBytes / 1MB, 2)
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