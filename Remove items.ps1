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

.PARAMETER OlderThanUnit
    Mandatory for 'File' and 'FilesInFolder'. Day, Month or Year.

.PARAMETER OlderThanQuantity
    Mandatory for 'File' and 'FilesInFolder'. Value 0 removes all files
    regardless of their creation date.

.PARAMETER Recurse
    Also remove the files in the subfolders, for 'FilesInFolder' only.
#>

Param (
    [Parameter(Mandatory)]
    [ValidateSet('File', 'FilesInFolder', 'EmptyFolders')]
    [String]$Type,
    [Parameter(Mandatory)]
    [String]$Path,
    [ValidateSet('Day', 'Month', 'Year')]
    [String]$OlderThanUnit,
    [Int]$OlderThanQuantity,
    [Boolean]$Recurse
)

if ($Type -eq 'EmptyFolders') {
    $failedFolderRemoval = @()

    while (
        $emptyFolders = Get-ChildItem -LiteralPath $Path -Directory -Recurse |
        Where-Object {
            ($_.GetFileSystemInfos().Count -eq 0) -and
            ($failedFolderRemoval -notContains $_.FullName)
        }
    ) {
        foreach ($emptyFolder in $emptyFolders) {
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
                $failedFolderRemoval += $emptyFolder.FullName
            }
            finally {
                $result
            }
        }
    }

    return
}

if (
    -not (
        $PSBoundParameters.ContainsKey('OlderThanUnit') -and
        $PSBoundParameters.ContainsKey('OlderThanQuantity')
    )
) {
    throw "Parameters 'OlderThanUnit' and 'OlderThanQuantity' are mandatory for type '$Type'"
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
$getErrors = @()

$files = if ($Type -eq 'File') {
    Get-Item -LiteralPath $Path -ErrorAction Stop
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
#endregion

#region Select files older than
Write-Verbose "Select files with a creation date older than '$OlderThanQuantity $OlderThanUnit'"

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
        { $_.CreationTime.ToString($dateFormat) -le $cutoff }
    )
}
#endregion

foreach ($fileToRemove in $files) {
    try {
        Write-Verbose "Remove file '$($fileToRemove.FullName)'"

        $result = [PSCustomObject]@{
            DateTime     = Get-Date
            ComputerName = $env:COMPUTERNAME
            Type         = 'File'
            FullName     = $fileToRemove.FullName
            CreationTime = $fileToRemove.CreationTime
            Action       = $null
            Error        = $null
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
