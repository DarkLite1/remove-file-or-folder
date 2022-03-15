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
            [ValidateSet('file', 'folder', 'content')]
            [String]$Type,
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [Int]$OlderThanDays,
            [Boolean]$RemoveEmptyFolders
        )

        Try {
            $compareDate = (Get-Date).AddDays(-$OlderThanDays)
    
            $result = [PSCustomObject]@{
                Type               = $Type
                Path               = $Path
                OlderThanDays      = $OlderThanDays
                OlderThanDate      = $compareDate
                ComputerName       = $env:COMPUTERNAME
                RemoveEmptyFolders = $RemoveEmptyFolders
                Items              = @()
                Error              = $null
            }

            #region Test file folder content
            if (
                ($Type -eq 'content') -and
                (-not (Test-Path -LiteralPath $Path -PathType Container))
            ) {
                throw 'Folder not found'
            }
            elseif (
                ($Type -eq 'folder') -and
                (-not (Test-Path -LiteralPath $Path -PathType Container))
            ) {
                $result.Items += [PSCustomObject]@{
                    Type         = 'Folder'
                    FullName     = $Path
                    CreationTime = $null
                    Action       = $null
                    Error        = 'Path not found'
                }
                Exit
            }
            elseif (
                ($Type -eq 'file') -and
                (-not (Test-Path -LiteralPath $Path -PathType Leaf))
            ) {
                $result.Items += [PSCustomObject]@{
                    Type         = 'File'
                    FullName     = $Path
                    CreationTime = $null
                    Action       = $null
                    Error        = 'Path not found'
                }
                Exit
            }
            #endregion

            #region Remove files
            $removeParams = @{
                Recurse     = $true
                Force       = $true
                ErrorAction = 'Stop'
            }
            $getParams = @{
                LiteralPath = $Path 
                Recurse     = $true
            }
            
            $result.Items = Get-ChildItem @getParams -File | 
            Where-Object { 
                ($_.CreationTime -lt $compareDate) -or
                ($OlderThanDays -eq 0)
            } | ForEach-Object {
                try {
                    Remove-Item @removeParams -LiteralPath $_.FullName
                    [PSCustomObject]@{
                        Type         = 'File' 
                        FullName     = $_.FullName 
                        CreationTime = $_.CreationTime
                        Action       = 'Removed'
                        Error        = $null
                    }
                }
                catch {
                    [PSCustomObject]@{
                        Type         = 'File' 
                        FullName     = $_.FullName 
                        CreationTime = $_.CreationTime
                        Action       = $null
                        Error        = $_
                    }
                    $Error.RemoveAt(0)
                }
            }
            #endregion

            #region Remove empty folders
            if (
                ($Type -eq 'content') -and
                ($RemoveEmptyFolders)
            ) {
                $failedFolderRemoval = @()

                while (
                    $emptyFolders = Get-ChildItem @getParams -Directory | 
                    Where-Object { 
                        ($_.GetFileSystemInfos().Count -eq 0) -and 
                        ($failedFolderRemoval -notContains $_.FullName) 
                    }
                ) {
                    $emptyFolders | ForEach-Object {
                        try {
                            Remove-Item @removeParams -LiteralPath $_.FullName
                            [PSCustomObject]@{
                                Type         = 'Folder' 
                                FullName     = $_.FullName 
                                CreationTime = $_.CreationTime
                                Action       = 'Removed'
                                Error        = $null
                            }
                        }
                        catch {
                            [PSCustomObject]@{
                                Type         = 'Folder' 
                                FullName     = $_.FullName 
                                CreationTime = $_.CreationTime
                                Action       = $null
                                Error        = $_
                            }
                            $Error.RemoveAt(0)
                            $failedFolderRemoval += $_.FullName
                        }
                    }   
                }
            }
            #endregion
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
        #endregion

        #region Test .json file properties
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
            if ($d.Remove -notMatch '^folder$|^file$|^content$') {
                throw "Input file '$ImportFile': Value '$($d.Remove)' in 'Remove' is not valid, only values 'folder', 'file' or 'content' are supported."
            }
            #endregion

            #region RemoveEmptyFolders
            if ($d.Remove -eq 'content') {
                if (
                    $d.PSObject.Properties.Name -notContains 'RemoveEmptyFolders'
                ) {
                    throw "Input file '$ImportFile': No 'RemoveEmptyFolders' found."
                }
                if (-not ($d.RemoveEmptyFolders -is [boolean])) {
                    throw "Input file '$ImportFile': The value '$($d.RemoveEmptyFolders)' in 'RemoveEmptyFolders' is not a true false value."
                }
            }
            else {
                if ($d.RemoveEmptyFolders) {
                    throw "Input file '$ImportFile': 'RemoveEmptyFolders' cannot be used with 'Remove' value '$($d.Remove)'."
                }
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
                ArgumentList = $d.Remove, $d.Path, $d.OlderThanDays
                AsJob        = $true
            }
            if (-not $d.ComputerName) { 
                $invokeParams.ComputerName = $env:COMPUTERNAME
            }
            if ($d.RemoveEmptyFolders) {
                $invokeParams.ArgumentList += $d.RemoveEmptyFolders
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
        $exportToExcel = foreach (
            $job in 
            $jobResults | Where-Object { $_.Items }
        ) {
            $job.Items | Select-Object -Property @{
                Name       = 'ComputerName'; 
                Expression = { $job.ComputerName } 
            },
            'Type', @{
                Name       = 'Path'; 
                Expression = { $_.FullName } 
            }, 'CreationTime', 'Action', 'Error'
        }

        if ($exportToExcel) {
            $M = "Export $($exportToExcel.Count) rows to Excel"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
            $excelParams = @{
                Path               = $LogFile + '- Log.xlsx'
                WorksheetName      = 'Overview'
                TableName          = 'Overview'
                NoNumberConversion = '*'
                AutoSize           = $true
                FreezeTopRow       = $true
            }
            $exportToExcel | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Send mail to user
        $removedItems = $jobResults.Where( { ($_.Action -eq 'Remove') })
        $removalErrors = $jobResults.Where( { $_.Error })
           
        $mailParams.Subject = "Removed $($removedItems.Count)/$($importExcelFile.count) items"

        $ErrorTable = $null
   
        # $error is polluted by VsCode
        if ($errorMessages = $Error.Exception.Message | Where-Object { $_ }) {
            $mailParams.Priority = 'High'
            $mailParams.Subject = "$($errorMessages.Count) errors, $($mailParams.Subject)"
            $ErrorTable = "<p>During removal <b>$($Error.Count) non terminating errors</b> were detected:$($errorMessages | ConvertTo-HtmlListHC)</p>"
        }

        if ($removalErrors) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $($removalErrors.Count) removal errors"
        }
   
        $table = "
           <table>
               <tr>
                   <th>Successfully removed items</th>
                   <td>$($removedItems.Count)</td>
               </tr>
               <tr>
                   <th>Errors while removing items</th>
                   <td>$($removalErrors.Count)</td>
               </tr>
               <tr>
                   <th>Not existing items after running the script</th>
                   <td>$($notExistingItems.Count)</td>
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