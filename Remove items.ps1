#Requires -Version 7

<#
.SYNOPSIS
    Remove a file, the files in a folder or the empty folders in a folder.

.DESCRIPTION
    Runs on the computer that holds the files, locally or through
    Invoke-Command -FilePath, so it cannot depend on other files.

    Returns one object per removed item or error.

.PARAMETER Type
    File          : remove the file in 'Path'
    FilesInFolder : remove the files in the folder 'Path'
    EmptyFolders  : remove all empty folders in the folder 'Path'

.PARAMETER ExcludeFolder
    Folders below 'Path' to skip, for 'FilesInFolder' and 'EmptyFolders'. The
    files and folders inside them are never removed.

.PARAMETER OlderThanUnit
    Mandatory for 'File' and 'FilesInFolder'. Day, Month or Year.

.PARAMETER OlderThanQuantity
    Mandatory for 'File' and 'FilesInFolder'. Value 0 removes all files
    regardless of their creation date.

.PARAMETER Recurse
    Also remove the files in the subfolders, for 'FilesInFolder' only.

.PARAMETER ExcludeFile
    Files below 'Path' that are never removed, for 'FilesInFolder' only.

.PARAMETER OlderThanBasedOn
    Mandatory for 'File' and 'FilesInFolder'. The file date compared with
    'OlderThan': CreationTime or LastWriteTime.
#>

param (
    [Parameter(Mandatory)]
    [ValidateSet('File', 'FilesInFolder', 'EmptyFolders')]
    [String]$Type,
    [Parameter(Mandatory)]
    [String]$Path,
    [AllowEmptyCollection()]
    [String[]]$ExcludeFolder = @(),
    [ValidateSet('Day', 'Month', 'Year')]
    [String]$OlderThanUnit,
    [ValidateRange(0, [int]::MaxValue)]
    [Int]$OlderThanQuantity,
    [Boolean]$Recurse,
    [AllowEmptyCollection()]
    [String[]]$ExcludeFile = @(),
    [ValidateSet('CreationTime', 'LastWriteTime')]
    [String]$OlderThanBasedOn
)

function Get-NormalizedPathHC {
    param ([String]$Value)

    [System.IO.Path]::GetFullPath(
        $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Value)
    )
}

$excludedPaths = @(
    $ExcludeFolder | Where-Object { $_ } | ForEach-Object { (Get-NormalizedPathHC $_).TrimEnd('\') }
)
$excludedFiles = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@($ExcludeFile | Where-Object { $_ } | ForEach-Object { Get-NormalizedPathHC $_ }),
    [StringComparer]::OrdinalIgnoreCase
)

function Test-IsExcludedHC {
    param ([String]$FullName)

    foreach ($excludedPath in $excludedPaths) {
        if (
            $FullName.Equals($excludedPath, [StringComparison]::OrdinalIgnoreCase) -or
            $FullName.StartsWith("$excludedPath\", [StringComparison]::OrdinalIgnoreCase)
        ) {
            return $true
        }
    }
    $false
}

function Get-ExclusiveCutoffHC {
    param (
        [Parameter(Mandatory)]
        [datetime]$ReferenceDate,
        [Parameter(Mandatory)]
        [ValidateSet('Day', 'Month', 'Year')]
        [string]$Unit,
        [Parameter(Mandatory)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Quantity
    )

    try {
        switch ($Unit) {
            'Day' { $ReferenceDate.AddDays(-$Quantity).Date.AddDays(1) }
            'Month' {
                $cutoffDate = $ReferenceDate.AddMonths(-$Quantity)
                [datetime]::new($cutoffDate.Year, $cutoffDate.Month, 1).AddMonths(1)
            }
            'Year' {
                $cutoffDate = $ReferenceDate.AddYears(-$Quantity)
                [datetime]::new($cutoffDate.Year, 1, 1).AddYears(1)
            }
        }
    }
    catch { throw "Invalid retention period '$Quantity $Unit': $_" }
}

function New-ReadErrorResultHC {
    param (
        [string]$FullName,
        [string]$ItemType,
        [string]$Message
    )

    [PSCustomObject]@{
        DateTime     = Get-Date
        ComputerName = $env:COMPUTERNAME
        Type         = $ItemType
        FullName     = $FullName
        CreationTime = $null
        Action       = $null
        Error        = $Message
    }
}

function Get-IncludedChildItemHC {
    param (
        [string]$Root,
        [bool]$Recursive,
        [switch]$Directories,
        [System.Collections.Generic.List[object]]$ReadErrors
    )

    $pending = [System.Collections.Generic.Stack[string]]::new()
    $rootPath = Get-NormalizedPathHC $Root
    if (Test-IsExcludedHC $rootPath) { return }
    $pending.Push($rootPath)
    $directoryFilter = if ($Directories) { @{ Directory = $true } } else { @{} }

    while ($pending.Count) {
        $directoryPath = $pending.Pop()
        $enumerationErrors = @()
        Get-ChildItem -LiteralPath $directoryPath -Force @directoryFilter -ErrorAction SilentlyContinue -ErrorVariable enumerationErrors |
        ForEach-Object {
            if ($_.PSIsContainer) {
                if (-not (Test-IsExcludedHC $_.FullName)) {
                    if ($Recursive -and -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                        $pending.Push($_.FullName)
                    }
                    if ($Directories) { $_ }
                }
            }
            elseif (-not $Directories) { $_ }
        }
        foreach ($readError in $enumerationErrors) { $ReadErrors.Add($readError) }
    }
}

if ($Type -eq 'EmptyFolders') {
    $getErrors = [System.Collections.Generic.List[object]]::new()
    $unreadableFolders = @{}

    function Test-IsEmptyFolderHC {
        param ([System.IO.DirectoryInfo]$Folder)

        $iterator = $null
        try {
            $iterator = $Folder.EnumerateFileSystemInfos().GetEnumerator()
            -not $iterator.MoveNext()
        }
        catch {
            $unreadableFolders[$Folder.FullName] = $_.Exception.InnerException.Message
            $Error.RemoveAt(0)
            $false
        }
        finally {
            if ($null -ne $iterator) { $iterator.Dispose() }
        }
    }

    $getParams = @{
        LiteralPath   = $Path
        Directory     = $true
        Recurse       = $true
        Force         = $true
        ErrorAction   = 'SilentlyContinue'
        ErrorVariable = '+getErrors'
    }

    $folderCandidates = if ($excludedPaths) {
        Get-IncludedChildItemHC -Root $Path -Recursive $true -Directories -ReadErrors $getErrors |
        Sort-Object { $_.FullName.Length } -Descending
    }
    else {
        Get-ChildItem @getParams | Sort-Object { $_.FullName.Length } -Descending
    }

    $folderCandidates | Where-Object { Test-IsEmptyFolderHC $_ } | ForEach-Object {
        $emptyFolder = $_
        try {
            Write-Verbose "Remove empty folder '$emptyFolder'"

            $result = [PSCustomObject]@{
                DateTime     = [datetime]::Now
                ComputerName = $env:COMPUTERNAME
                Type         = 'EmptyFolder'
                FullName     = $emptyFolder.FullName
                CreationTime = $emptyFolder.CreationTime
                Action       = $null
                Error        = $null
            }

            # non-recursive delete fails when the folder is no longer empty
            if ($emptyFolder.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                $emptyFolder.Attributes = $emptyFolder.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly
            }
            $emptyFolder.Delete()
            $result.Action = 'Removed'
        }
        catch {
            Write-Verbose "Failed to remove empty folder '$emptyFolder': $_"

            $result.Error = $_
            $Error.RemoveAt(0)
        }
        finally {
            $result
        }
    }

    $readErrors = @{}

    foreach ($getError in $getErrors) {
        $readErrors["$($getError.TargetObject)"] = "$getError"
    }
    foreach ($unreadableFolder in $unreadableFolders.GetEnumerator()) {
        $readErrors[$unreadableFolder.Key] = $unreadableFolder.Value
    }

    $readErrors.GetEnumerator() |
    Where-Object { -not (Test-IsExcludedHC $_.Key) } |
    Sort-Object -Property Key |
    ForEach-Object {
        Write-Warning "Failed to read '$($_.Key)': $($_.Value)"

        New-ReadErrorResultHC -FullName $_.Key -ItemType $Type -Message $_.Value
    }

    return
}

if (
    -not (
        $PSBoundParameters.ContainsKey('OlderThanUnit') -and
        $PSBoundParameters.ContainsKey('OlderThanQuantity') -and
        $PSBoundParameters.ContainsKey('OlderThanBasedOn')
    )
) {
    throw "Parameters 'OlderThanUnit', 'OlderThanQuantity' and 'OlderThanBasedOn' are mandatory for type '$Type'"
}

#region Test path exists
$pathType = if ($Type -eq 'File') { 'Leaf' } else { 'Container' }

if (-not (Test-Path -LiteralPath $Path -PathType $pathType)) {
    return New-ReadErrorResultHC -FullName $Path -ItemType $Type -Message 'Path not found'
}
#endregion

#region Select files older than
Write-Verbose "Select files with a $OlderThanBasedOn older than '$OlderThanQuantity $OlderThanUnit'"

if ($OlderThanQuantity -ne 0) {
    $cutoffExclusive = Get-ExclusiveCutoffHC -ReferenceDate (Get-Date) -Unit $OlderThanUnit -Quantity $OlderThanQuantity
}
#endregion

$getErrors = [System.Collections.Generic.List[object]]::new()

& {
    $fileReadErrors = @()
    if ($Type -eq 'File') {
        Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue -ErrorVariable fileReadErrors
    }
    elseif ($excludedPaths) {
        Get-IncludedChildItemHC -Root $Path -Recursive $Recurse -ReadErrors $getErrors
    }
    else {
        Get-ChildItem -LiteralPath $Path -File -Recurse:$Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable fileReadErrors
    }
    foreach ($readError in $fileReadErrors) { $getErrors.Add($readError) }
} | ForEach-Object {
    $fileToRemove = $_
    if ($excludedFiles.Contains($fileToRemove.FullName)) { return }
    if (($Type -eq 'File') -and $excludedPaths -and (Test-IsExcludedHC $fileToRemove.FullName)) { return }
    if (($OlderThanQuantity -ne 0) -and ($fileToRemove.$OlderThanBasedOn -ge $cutoffExclusive)) { return }

    try {
        Write-Verbose "Remove file '$($fileToRemove.FullName)'"

        $result = [PSCustomObject]@{
            DateTime      = [datetime]::Now
            ComputerName  = $env:COMPUTERNAME
            Type          = 'File'
            FullName      = $fileToRemove.FullName
            CreationTime  = $fileToRemove.CreationTime
            LastWriteTime = $fileToRemove.LastWriteTime
            Action        = $null
            Error         = $null
        }

        $fileToRemove.Refresh()
        if (-not $fileToRemove.Exists) {
            throw [System.IO.FileNotFoundException]::new('File no longer exists', $fileToRemove.FullName)
        }
        if (($OlderThanQuantity -ne 0) -and ($fileToRemove.$OlderThanBasedOn -ge $cutoffExclusive)) {
            $result = $null
            return
        }
        $result.CreationTime = $fileToRemove.CreationTime
        $result.LastWriteTime = $fileToRemove.LastWriteTime
        if ($fileToRemove.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
            $fileToRemove.Attributes = $fileToRemove.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly
        }
        $fileToRemove.Delete()

        $result.Action = 'Removed'
    }
    catch {
        Write-Warning "Failed to remove file '$($result.FullName)': $_"

        $result.Error = $_
        $Error.RemoveAt(0)
    }
    finally {
        if ($null -ne $result) { $result }
    }
}

foreach ($getError in $getErrors) {
    if ($excludedPaths -and (Test-IsExcludedHC "$($getError.TargetObject)")) { continue }
    Write-Warning "Failed to read '$($getError.TargetObject)': $getError"

    New-ReadErrorResultHC -FullName "$($getError.TargetObject)" -ItemType $Type -Message "$getError"
}
