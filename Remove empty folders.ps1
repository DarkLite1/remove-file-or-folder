Param (
    [Parameter(Mandatory)]
    [String]$Path
)

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