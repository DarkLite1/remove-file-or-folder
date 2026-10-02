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

Param (
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
$excludedFiles = @(
    $ExcludeFile | Where-Object { $_ } | ForEach-Object { Get-NormalizedPathHC $_ }
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

    while ($pending.Count) {
        $directoryPath = $pending.Pop()
        $enumerationErrors = @()
        Get-ChildItem -LiteralPath $directoryPath -Force -ErrorAction SilentlyContinue -ErrorVariable enumerationErrors |
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

        try {
            $Folder.GetFileSystemInfos().Count -eq 0
        }
        catch {
            $unreadableFolders[$Folder.FullName] = $_.Exception.InnerException.Message
            $Error.RemoveAt(0)
            $false
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

    $emptyFolders = if ($excludedPaths) {
        Get-IncludedChildItemHC -Root $Path -Recursive $true -Directories -ReadErrors $getErrors |
        Sort-Object { $_.FullName.Length } -Descending
    }
    else {
        Get-ChildItem @getParams | Sort-Object { $_.FullName.Length } -Descending
    }

    $emptyFolders | Where-Object { Test-IsEmptyFolderHC $_ } | ForEach-Object {
        $emptyFolder = $_
        try {
            Write-Verbose "Remove empty folder '$emptyFolder'"

            $result = [PSCustomObject]@{
                DateTime     = Get-Date
                ComputerName = $env:COMPUTERNAME
                Type         = 'EmptyFolder'
                FullName     = $emptyFolder.FullName
                CreationTime = $emptyFolder.CreationTime
                Action       = $null
                Error        = $null
            }

            # non-recursive delete fails when the folder is no longer empty
            $emptyFolder.Attributes = $emptyFolder.Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly
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

        [PSCustomObject]@{
            DateTime     = Get-Date
            ComputerName = $env:COMPUTERNAME
            Type         = $Type
            FullName     = $_.Key
            CreationTime = $null
            Action       = $null
            Error        = $_.Value
        }
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
    return [PSCustomObject]@{
        DateTime     = Get-Date
        ComputerName = $env:COMPUTERNAME
        Type         = $Type
        FullName     = $Path
        CreationTime = $null
        Action       = $null
        Error        = 'Path not found'
    }
}
#endregion

#region Get files
$getErrors = [System.Collections.Generic.List[object]]::new()

$files = if ($Type -eq 'File') {
    Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue -ErrorVariable getErrors
}
elseif ($excludedPaths) {
    Get-IncludedChildItemHC -Root $Path -Recursive $Recurse -ReadErrors $getErrors
}
else {
    $getParams = @{
        LiteralPath   = $Path
        Recurse       = $Recurse
        File          = $true
        Force         = $true
        ErrorAction   = 'SilentlyContinue'
        ErrorVariable = 'getErrors'
    }
    Get-ChildItem @getParams
}

if ($excludedPaths -and ($Type -eq 'File')) {
    $files = $files.Where({ -not (Test-IsExcludedHC $_.FullName) })
    $getErrors = $getErrors.Where({ -not (Test-IsExcludedHC "$($_.TargetObject)") })
}

if ($excludedFiles) {
    $files = $files.Where({
            $fullName = $_.FullName
            -not $excludedFiles.Where({
                    $_.Equals($fullName, [StringComparison]::OrdinalIgnoreCase)
                })
        })
}
#endregion

#region Select files older than
Write-Verbose "Select files with a $OlderThanBasedOn older than '$OlderThanQuantity $OlderThanUnit'"

if ($OlderThanQuantity -ne 0) {
    $today = Get-Date

    # compares calendar periods: 'older than 1 month' is any earlier month
    $dateFormat, $cutoffDate = switch ($OlderThanUnit) {
        'Day' { 'yyyyMMdd', $today.AddDays(-$OlderThanQuantity) }
        'Month' { 'yyyyMM', $today.AddMonths(-$OlderThanQuantity) }
        'Year' { 'yyyy', $today.AddYears(-$OlderThanQuantity) }
    }
    $cutoff = $cutoffDate.ToString($dateFormat)

    $files = $files.Where(
        { $_.$OlderThanBasedOn.ToString($dateFormat) -le $cutoff }
    )
}
#endregion

foreach ($fileToRemove in $files) {
    try {
        Write-Verbose "Remove file '$($fileToRemove.FullName)'"

        $result = [PSCustomObject]@{
            DateTime      = Get-Date
            ComputerName  = $env:COMPUTERNAME
            Type          = 'File'
            FullName      = $fileToRemove.FullName
            CreationTime  = $fileToRemove.CreationTime
            LastWriteTime = $fileToRemove.LastWriteTime
            Action        = $null
            Error         = $null
        }

        $params = @{
            LiteralPath = $fileToRemove.FullName
            Force       = $true
            ErrorAction = 'Stop'
        }
        Remove-Item @params

        $result.Action = 'Removed'
    }
    catch {
        Write-Warning "Failed to remove file '$($result.FullName)': $_"

        $result.Error = $_
        $Error.RemoveAt(0)
    }
    finally {
        $result
    }
}

foreach ($getError in $getErrors) {
    Write-Warning "Failed to read '$($getError.TargetObject)': $getError"

    [PSCustomObject]@{
        DateTime     = Get-Date
        ComputerName = $env:COMPUTERNAME
        Type         = $Type
        FullName     = "$($getError.TargetObject)"
        CreationTime = $null
        Action       = $null
        Error        = "$getError"
    }
}
