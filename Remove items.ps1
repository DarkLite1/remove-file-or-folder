#Requires -Version 7

<#
.SYNOPSIS
    Remove a file, the files in a folder or the empty folders in a folder.

.DESCRIPTION
    Runs on the computer that holds the files, locally or through
    Invoke-Command -FilePath, so it cannot depend on other files.

    File age is based on CreationTime or LastWriteTime. Day, Month and Year
    include the entire cutoff day, month or year, plus earlier dates. These
    are calendar cutoffs, not elapsed time. For example, 1 Day at 00:05 can
    select a file written at 23:55 yesterday. Quantity 0 disables this filter.

    Hidden, system and read-only items are included by default. ExcludeAttributes
    can protect Hidden or System items. Excluded folder trees are not traversed.
    Directory links encountered during traversal are not followed.
    Empty subfolders are processed deepest first; the root is never removed.

    Deletes items immediately; there is no preview or WhatIf mode. Main.ps1
    normally calls this worker after validating the JSON configuration.

.PARAMETER Type
    File          : remove the file in Path when it passes the age filter.
    FilesInFolder : remove matching files below Path, honoring Recurse.
    EmptyFolders  : remove empty subfolders at every depth below Path.

.PARAMETER Path
    Literal file path for File, or root folder path for FilesInFolder and
    EmptyFolders. Wildcards are not expanded. Relative paths are resolved
    on the computer executing the worker.

.PARAMETER ExcludeFolder
    Folder paths to protect, including their contents. Paths are normalized
    and compared without regard to case. For folder jobs, excluded trees
    are skipped. For File, a file inside an excluded folder is also skipped.

.PARAMETER OlderThanUnit
    Required for File and FilesInFolder, even when Quantity is 0. Day, Month
    or Year. Ignored for EmptyFolders; folder cleanup is not age-based.

.PARAMETER OlderThanQuantity
    Required for File and FilesInFolder. Whole number of 0 or greater.
    Value 0 removes selected files regardless of the date in OlderThanBasedOn;
    exclusions and Recurse still apply. Ignored for EmptyFolders.

.PARAMETER Recurse
    Boolean for FilesInFolder. Pass $true to include files in subfolders,
    or $false (the default) for files directly in Path. EmptyFolders always
    checks every depth, regardless of this value.

.PARAMETER ExcludeFile
    Exact file paths to protect in file jobs. Paths are normalized and
    compared without regard to case; wildcards are not supported. Ignored
    for EmptyFolders, which never deletes folders that contain files.

.PARAMETER OlderThanBasedOn
    Required for File and FilesInFolder, even when Quantity is 0. The file
    timestamp to compare: CreationTime or LastWriteTime. Ignored for
    EmptyFolders.

.PARAMETER ExcludeAttributes
    Optional Hidden and/or System attributes. An item matching either is
    skipped. Matching folders, including a matching root, are not entered.
    For File, checks the selected file's own attributes. An empty list keeps
    the default behavior. Does not suppress genuine access errors.

.OUTPUTS
    PSCustomObject. One result per removal or error, not per skipped item.
    Results contain DateTime, ComputerName, Type, FullName, CreationTime,
    Action and Error. File-removal results also contain LastWriteTime.
    Action is 'Removed' on success; Error describes failures. Invalid
    parameters or retention periods can throw instead of returning a result.

.EXAMPLE
    & '.\Remove items.ps1' -Type FilesInFolder -Path 'C:\Logs' -OlderThanUnit Day -OlderThanQuantity 30 -OlderThanBasedOn LastWriteTime -Recurse $true

    Removes files last written on or before the date 30 calendar days ago,
    including files in subfolders. Does not delete any folders.

.EXAMPLE
    & '.\Remove items.ps1' -Type EmptyFolders -Path 'C:\Drop'

    Removes empty subfolders at every depth, but keeps C:\Drop itself.
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
    [String]$OlderThanBasedOn,
    [AllowEmptyCollection()]
    [ValidateSet('Hidden', 'System')]
    [String[]]$ExcludeAttributes = @()
)

$excludedAttributeMask = [System.IO.FileAttributes]0
foreach ($attribute in $ExcludeAttributes) {
    $excludedAttributeMask = $excludedAttributeMask -bor [System.IO.FileAttributes]$attribute
}

function Get-NormalizedPathHC {
    <#
    .SYNOPSIS
        Resolve a PowerShell path to a full filesystem path without requiring
        the item to exist. Relative paths use the current working directory.
    #>
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
    <#
    .SYNOPSIS
        Test whether a full path is an excluded folder or is inside one.

    .DESCRIPTION
        Uses the worker's normalized excludedPaths list. Matching is
        case-insensitive and respects folder boundaries, so excluding
        C:\Log does not exclude C:\Logs.
    #>
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
    <#
    .SYNOPSIS
        Return the first timestamp that must be kept by a calendar age filter.

    .DESCRIPTION
        File timestamps strictly before the returned cutoff are eligible.
        The cutoff is midnight after the selected day, month or year.
        Throws when the requested date is outside the DateTime range.

    .PARAMETER ReferenceDate
        The current date and time on the computer executing the worker.

    .PARAMETER Unit
        Calendar unit to subtract: Day, Month or Year.

    .PARAMETER Quantity
        Positive whole number of units. The caller handles Quantity 0 by
        skipping age filtering instead of calling this helper.

    .EXAMPLE
        Get-ExclusiveCutoffHC -ReferenceDate ([datetime]'2026-10-02T12:00:00') -Unit Day -Quantity 1

        Returns 2 October 2026 at midnight, making every timestamp on
        1 October or earlier eligible for removal.
    #>
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
    <#
    .SYNOPSIS
        Create a result for an item that could not be found or read.

    .DESCRIPTION
        FullName identifies the failed path, ItemType identifies the worker
        operation, and Message becomes Error. CreationTime and Action are
        null. This helper returns data only; the caller writes any warning.
    #>
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
    <#
    .SYNOPSIS
        Enumerate files or folders without entering excluded folder trees.

    .DESCRIPTION
        Uses the worker's excludedPaths and excludedAttributeMask. Includes
        hidden items unless excluded, but does not recurse into directory
        links. Matching attribute-excluded folders are not entered, including
        the root. The root is never returned.

    .PARAMETER Root
        Folder to enumerate. An excluded root produces no candidates.

    .PARAMETER Recursive
        Include descendants when true; otherwise inspect only Root's children.

    .PARAMETER Directories
        Return directories instead of files.

    .PARAMETER ReadErrors
        Caller-owned list to which enumeration errors are added. Errors are
        kept separate from returned candidates for later reporting.
    #>
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
        if ($excludedAttributeMask) {
            try {
                if ([System.IO.File]::GetAttributes($directoryPath) -band $excludedAttributeMask) { continue }
            }
            catch {
                $ReadErrors.Add([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception, $_.FullyQualifiedErrorId, $_.CategoryInfo.Category, $directoryPath
                ))
                $Error.RemoveAt(0)
                continue
            }
        }
        $enumerationErrors = @()
        Get-ChildItem -LiteralPath $directoryPath -Force @directoryFilter -ErrorAction SilentlyContinue -ErrorVariable enumerationErrors |
        ForEach-Object {
            if ($_.Attributes -band $excludedAttributeMask) { return }
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

$writeVerbose = $VerbosePreference -notin 'SilentlyContinue', 'Ignore'

if ($Type -eq 'EmptyFolders') {
    $getErrors = [System.Collections.Generic.List[object]]::new()
    $unreadableFolders = @{}

    function Test-IsMissingSubfolderHC {
        <#
        .SYNOPSIS
            Confirm a failed subfolder is gone while its task root is accessible.

        .DESCRIPTION
            Existing paths, root paths and inconclusive access checks return
            false so their original errors remain visible.
        #>
        param ([string]$FullName)

        try {
            $rootPath = (Get-NormalizedPathHC $Path).TrimEnd('\')
            $folderPath = Get-NormalizedPathHC $FullName
            if (-not $folderPath.StartsWith("$rootPath\", [StringComparison]::OrdinalIgnoreCase)) {
                return $false
            }

            try {
                $null = [System.IO.File]::GetAttributes($folderPath)
                return $false
            }
            catch [System.IO.FileNotFoundException], [System.IO.DirectoryNotFoundException] {
                $Error.RemoveAt(0)
            }

            $rootAttributes = [System.IO.File]::GetAttributes($rootPath)
            return [bool]($rootAttributes -band [System.IO.FileAttributes]::Directory)
        }
        catch {
            $Error.RemoveAt(0)
            return $false
        }
    }

    function Test-IsEmptyFolderHC {
        <#
        .SYNOPSIS
            Check whether a folder has no entries, including hidden entries.

        .DESCRIPTION
            Reads at most one entry and disposes the enumerator. An unreadable
            folder returns false and is recorded in the worker's
            unreadableFolders table for later error reporting.
        #>
        param ([System.IO.DirectoryInfo]$Folder)

        $iterator = $null
        try {
            if ($excludedAttributeMask) {
                $Folder.Refresh()
                if ($Folder.Exists -and ($Folder.Attributes -band $excludedAttributeMask)) { return $false }
            }
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

    $folderCandidates = if ($excludedPaths -or $excludedAttributeMask) {
        Get-IncludedChildItemHC -Root $Path -Recursive $true -Directories -ReadErrors $getErrors |
        Sort-Object { $_.FullName.Length } -Descending
    }
    else {
        Get-ChildItem @getParams | Sort-Object { $_.FullName.Length } -Descending
    }

    $folderCandidates | Where-Object { Test-IsEmptyFolderHC $_ } | ForEach-Object {
        $emptyFolder = $_
        try {
            $result = [PSCustomObject]@{
                DateTime     = [datetime]::Now
                ComputerName = $env:COMPUTERNAME
                Type         = 'EmptyFolder'
                FullName     = $emptyFolder.FullName
                CreationTime = $emptyFolder.CreationTime
                Action       = $null
                Error        = $null
            }

            if ($excludedAttributeMask) {
                $emptyFolder.Refresh()
                if ($emptyFolder.Exists -and ($emptyFolder.Attributes -band $excludedAttributeMask)) {
                    $result = $null
                    return
                }
            }
            # non-recursive delete fails when the folder is no longer empty
            if ($emptyFolder.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                $emptyFolder.Attributes = $emptyFolder.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly
            }
            $emptyFolder.Delete()
            $result.Action = 'Removed'
            if ($writeVerbose) { Write-Verbose "Removed empty folder '$($emptyFolder.FullName)'" }
        }
        catch {
            if (Test-IsMissingSubfolderHC $emptyFolder.FullName) {
                $result = $null
            }
            else {
                Write-Warning "Failed to remove empty folder '$($emptyFolder.FullName)': $_"
                $result.Error = $_
            }
            $Error.RemoveAt(0)
        }
        finally {
            if ($null -ne $result) { $result }
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
        if (Test-IsMissingSubfolderHC $_.Key) { return }
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
if ($OlderThanQuantity -ne 0) {
    $cutoffExclusive = Get-ExclusiveCutoffHC -ReferenceDate (Get-Date) -Unit $OlderThanUnit -Quantity $OlderThanQuantity
    if ($writeVerbose) {
        Write-Verbose "Select files in '$Path' with $OlderThanBasedOn before $($cutoffExclusive.ToString('yyyy-MM-dd HH:mm:ss')) (exclusive calendar cutoff, local time)"
    }
}
elseif ($writeVerbose) {
    Write-Verbose "Age filtering disabled for '$Path'; exclusions and Recurse still apply"
}
#endregion

$getErrors = [System.Collections.Generic.List[object]]::new()

& {
    $fileReadErrors = @()
    if ($Type -eq 'File') {
        Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue -ErrorVariable fileReadErrors
    }
    elseif ($excludedPaths -or $excludedAttributeMask) {
        Get-IncludedChildItemHC -Root $Path -Recursive $Recurse -ReadErrors $getErrors
    }
    else {
        Get-ChildItem -LiteralPath $Path -File -Recurse:$Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable fileReadErrors
    }
    foreach ($readError in $fileReadErrors) { $getErrors.Add($readError) }
} | ForEach-Object {
    $fileToRemove = $_
    if ($fileToRemove.Attributes -band $excludedAttributeMask) { return }
    if ($excludedFiles.Contains($fileToRemove.FullName)) { return }
    if (($Type -eq 'File') -and $excludedPaths -and (Test-IsExcludedHC $fileToRemove.FullName)) { return }
    if (($OlderThanQuantity -ne 0) -and ($fileToRemove.$OlderThanBasedOn -ge $cutoffExclusive)) { return }

    try {
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
        if ($fileToRemove.Attributes -band $excludedAttributeMask) {
            $result = $null
            return
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
        if ($writeVerbose) { Write-Verbose "Removed file '$($fileToRemove.FullName)'" }
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
