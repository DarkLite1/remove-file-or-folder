#Requires -Version 5.1
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Remove files or folders on remote machine.

.DESCRIPTION
    The script reads an Excel file containing a computer name and a local folder
    or file path in each row. It then tries to remove the files or folders 
    defined on the requested computers.

.PARAMETER Path
    Path to the Excel file containing the rows with the computer names and local
    folder/file paths.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Remove file or folder\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    $scriptBlock = {
        Param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [Int]$OlderThanDays,
            [Parameter(Mandatory)]
            [Boolean]$RemoveEmptyFolders
        )

        Try {
            $result = [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                Path         = $Path
                Date         = Get-Date
                Exist        = $true
                Action       = $null
                Error        = $null
            }

            if (-not (Test-Path -LiteralPath $path)) {
                $result.Exist = $false
                $result.Error = 'Path not found'
                Continue
            }

            $result.Action = 'Remove'
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $path)) {
                $result.Exist = $false
            }
        }
        Catch {
            $result.Error = $_
            $Error.RemoveAt(0)
        }
        finally {
            $result
        }
    }

    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $LogParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $LogFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json

        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }
        if (-not ($Destinations = $file.Destinations)) {
            throw "Input file '$ImportFile': No 'Destinations' found."
        }
        foreach ($d in $Destinations) {
            #region Path
            if (-not $d.Path) {
                throw "Input file '$ImportFile': No 'Path' found in one of the 'Destinations'."
            }
            if (($d.Path -notMatch '^\\\\') -and (-not $d.ComputerName)) {
                throw "Input file '$ImportFile': No 'ComputerName' found for path '$($d.Path)' in 'Destinations'."
            }
            #endregion

            #region OlderThanDays
            if ($d.PSObject.Properties.Name -notContains 'OlderThanDays') {
                throw "Input file '$ImportFile': No 'OlderThanDays' number found. Number '0' removes all."
            }
            if (-not ($d.OlderThanDays -is [int])) {
                throw "Input file '$ImportFile': 'OlderThanDays' needs to be a number, the value '$($d.OlderThanDays)' is not supported. Use number '0' to remove all."
            }
            #endregion

            #region Remove
            if ($d.PSObject.Properties.Name -notContains 'Remove') {
                throw "Input file '$ImportFile': Property 'Remove' not found. Valid values are 'folder', 'file' or 'content'."
            }
            if ($d.Remove -notMatch 'folder|file|content') {
                throw "Input file '$ImportFile': Value '$($d.Remove)' in 'Remove' is not valid, only values 'folder', 'file' or 'content' are supported."
            }
            #endregion

            #region RemoveEmptyFolders
            if (
                ($d.Remove -eq 'content') -and
                ($d.PSObject.Properties.Name -notContains 'RemoveEmptyFolders')
            ) {
                throw "Input file '$ImportFile': No 'RemoveEmptyFolders' found."
            }
            if (-not ($d.RemoveEmptyFolders -is [boolean])) {
                throw "Input file '$ImportFile': The value '$($d.RemoveEmptyFolders)' in 'RemoveEmptyFolders' is not a true false value."
            }
            if (
                ($d.Remove -ne 'content') -and
                ($d.RemoveEmptyFolders)
            ) {
                throw "Input file '$ImportFile': 'RemoveEmptyFolders' cannot be used with 'Remove' value '$($d.Remove)'."
            }
            #endregion
        }
        #endregion

        $mailParams = @{ }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Remove files/folders on remote machines
        $jobs = foreach ($d in $Destinations) {
            $invokeParams = @{
                ComputerName = $d.ComputerName
                ScriptBlock  = $scriptBlock
                ArgumentList = $d.Path, $d.OlderThanDays, $d.RemoveEmptyFolders
                AsJob        = $true
            }
            
            if (-not $d.ComputerName) { 
                $invokeParams.ComputerName = $env:COMPUTERNAME
            }

            $M = "Start job on '{0}' for path '{1}' OlderThanDays '{2}' RemoveEmptyFolders '{3}'" -f $invokeParams.ComputerName,
            $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
            $invokeParams.ArgumentList[2]
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            Invoke-Command @invokeParams
        }

        $M = "Wait for all $($jobs.count) jobs to be finished"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $jobResults = if ($jobs) { $jobs | Wait-Job | Receive-Job }
        #endregion

        #region Export results to Excel log file
        if ($jobResults) {
            $M = "Export $($jobResults.Count) rows to Excel"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
            $excelParams = @{
                Path               = $LogFile + '- Log.xlsx'
                AutoSize           = $true
                WorksheetName      = 'Overview'
                TableName          = 'Overview'
                FreezeTopRow       = $true
                NoNumberConversion = '*'
            }
            $jobResults | 
            Select-Object -Property 'ComputerName', 'Path', 'Date', 
            'Exist', 'Action', 'Error' | 
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Send mail to user
        $removedFolders = $jobResults.Where( {
                ($_.Action -eq 'Remove') -and
                ($_.Exist -eq $false)
            })
        $folderRemovalErrors = $jobResults.Where( { $_.Error })
        $notExistingFolders = $jobResults.Where( { $_.Exist -eq $false })
           
        $mailParams.Subject = "Removed $($removedFolders.Count)/$($importExcelFile.count) items"

        $ErrorTable = $null
   
        if ($Error) {
            $mailParams.Priority = 'High'
            $mailParams.Subject = "$($Error.Count) errors, $($mailParams.Subject)"
            $ErrorTable = "<p>During removal <b>$($Error.Count) non terminating errors</b> were detected:$($Error.Exception | Select-Object -ExpandProperty Message | ConvertTo-HtmlListHC)</p>"
        }

        if ($folderRemovalErrors) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $($folderRemovalErrors.Count) removal errors"
        }
   
        $table = "
           <table>
               <tr>
                   <th>Successfully removed items</th>
                   <td>$($removedFolders.Count)</td>
               </tr>
               <tr>
                   <th>Errors while removing items</th>
                   <td>$($folderRemovalErrors.Count)</td>
               </tr>
               <tr>
                   <th>Imported Excel file rows</th>
                   <td>$($importExcelFile.Count)</td>
               </tr>
               <tr>
                   <th>Not existing items after running the script</th>
                   <td>$($notExistingFolders.Count)</td>
               </tr>
           </table>
           "
   
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "<p>Summary of removed items (files or folders):</p>
                $table
                $ErrorTable
                <p><i>* Check the attachment for details</i></p>"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }
   
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Get-Job | Remove-Job
        Write-EventLog @EventEndParams
    }
}