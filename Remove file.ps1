Param (
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [ValidateSet('Day', 'Month', 'Year')]
    [String]$OlderThanUnit,
    [Parameter(Mandatory)]
    [Int]$OlderThanQuantity
)

#region Test file exists
if (-not (Test-Path -LiteralPath $Path -PathType 'Leaf')) {
    return [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        Type         = 'File'
        FullName     = $null
        CreationTime = $null
        Action       = $null
        Error        = 'Path not found'
    }
}
#endregion

#region Create filter
Write-Verbose "Create filter for files with a creation date older than '$OlderThanQuantity $OlderThanUnit'"

if ($OlderThanQuantity -eq 0) {
    Filter Select-FileHC {
        Write-Output $_
    }
}
else {
    $today = Get-Date

    Switch ($OlderThanUnit) {
        'Day' {
            Filter Select-FileHC {
                if (
                    $_.CreationTime.Date.ToString('yyyyMMdd') -le $(($today.AddDays( - $OlderThanQuantity)).Date.ToString('yyyyMMdd'))
                ) {
                    Write-Output $_
                }
            }

            break
        }
        'Month' {
            Filter Select-FileHC {
                if (
                    $_.CreationTime.Date.ToString('yyyyMM') -le $(($today.AddMonths( - $OlderThanQuantity)).Date.ToString('yyyyMM'))
                ) {
                    Write-Output $_
                }
            }

            break
        }
        'Year' {
            Filter Select-FileHC {
                if (
                    $_.CreationTime.Date.ToString('yyyy') -le $(($today.AddYears( - $OlderThanQuantity)).Date.ToString('yyyy'))
                ) {
                    Write-Output $_
                }
            }

            break
        }
        Default {
            throw "OlderThan.Unit '$_' not supported"
        }
    }
}
#endregion

Get-Item -LiteralPath $Path -ErrorAction Stop | Select-FileHC | ForEach-Object {
    try {
        Write-Verbose "Remove file '$Path'"

        $result = [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            Type         = 'File'
            FullName     = $_.FullName
            CreationTime = $_.CreationTime
            Action       = $null
            Error        = $null
        }

        $params = @{
            LiteralPath = $_.FullName
            Force       = $true
            ErrorAction = 'Stop'
        }
        Remove-Item @params

        $result.Action = 'Removed'
    }
    catch {
        Write-Warning "Failed to remove file '$Path': $_"

        $result.Error = $_
        $Error.RemoveAt(0)
    }
    finally {
        $result
    }
}