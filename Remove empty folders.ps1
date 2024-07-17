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
                ComputerName = $env:COMPUTERNAME
                Type         = 'EmptyFolder'
                FullName     = $emptyFolder.FullName
                CreationTime = $emptyFolder.CreationTime
                Action       = $null
                Error        = $null
            }

            $params = @{
                LiteralPath = $emptyFolder.FullName
                Recurse     = $true
                Force       = $true
                ErrorAction = 'Stop'
            }
            Remove-Item @params
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