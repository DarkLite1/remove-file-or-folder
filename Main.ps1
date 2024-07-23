#Requires -Version 7
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.HTML, Toolbox.Remoting

<#
.SYNOPSIS
    Remove files, folders or folder content on remote machines.

.DESCRIPTION
    This script reads a .JSON file containing the destination paths where files
    or folders need to be removed. When 'OlderThanDays' is '0' all files or
    folders will be removed, depending on the chosen 'Remove' type, regardless
    their creation date.

.PARAMETER Remove.File
    A collection of file paths that need to be removed.

.PARAMETER Remove.FilesInFolder
    A collection of folder paths where to look for files that need to be
    removed.

.PARAMETER Remove.EmptyFolders
    A collection of folder paths where to look for empty folders that need to be
    removed.

.PARAMETER Remove.File.Name
    The name to display in the email send to the user instead of the full path.

.PARAMETER Remove.File.ComputerName
    Computer name where the removal action will be executed.

.PARAMETER Remove.File.Path
    Can be a local path when 'ComputerName' is used or a UNC path

.PARAMETER Remove.File.OlderThan.Quantity
    Number of units (day, ..).

.PARAMETER Remove.File.OlderThan.Unit
    Valid options:
    - Day
    - Month
    - Year

.PARAMETER MaxConcurrentJobs
    Determines how many jobs to run at the same time

.PARAMETER SendMail.To
    List of e-mail addresses where to send the e-mail too.

.PARAMETER SendMail.When
    When to send an e-mail.

    Valid options:
    - Never               : Never send an e-mail
    - Always              : Always send an e-mail, even without matches
    - OnlyOnErrorOrAction : Only send an e-mail when action is taken or on error
    - OnlyOnError         : Only send an e-mail on error

.PARAMETER PSSessionConfiguration
    The version of PowerShell on the remote endpoint as returned by
    Get-PSSessionConfiguration.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$PSSessionConfiguration = 'PowerShell.7',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Remove file or folder\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        $Error.Clear()

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @logParams
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
        try {
            @(
                'SendMail', 'MaxConcurrentJobs', 'Remove'
            ).where(
                { -not $file.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            @(
                'To', 'When'
            ).where(
                { -not $file.SendMail.$_ }
            ).foreach(
                { throw "Property 'SendMail.$_' not found" }
            )

            if ($file.SendMail.When -notMatch '^Never$|^OnlyOnError$|^OnlyOnErrorOrAction$|^Always$') {
                throw "Value '$($file.SendMail.When)' for 'SendMail.When' is not supported. Supported values are 'Never, OnlyOnError, OnlyOnErrorOrAction or Always'"
            }

            $MaxConcurrentJobs = $file.MaxConcurrentJobs
            try {
                $null = $MaxConcurrentJobs.ToInt16($null)
            }
            catch {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$MaxConcurrentJobs' is not supported."
            }

            foreach ($fileToRemove in $file.Remove.File) {
                @(
                    'Path', 'OlderThan'
                ).where(
                    { -not $fileToRemove.$_ }
                ).foreach(
                    { throw "Property 'Remove.File.$_' not found" }
                )

                #region OlderThan
                if (-not $fileToRemove.OlderThan.Unit) {
                    throw "No 'Remove.File.OlderThan.Unit' found"
                }

                if ($fileToRemove.OlderThan.Unit -notMatch '^Day$|^Month$|^Year$') {
                    throw "Value '$($fileToRemove.OlderThan.Unit)' is not supported by 'Remove.File.OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'."
                }

                if ($fileToRemove.OlderThan.PSObject.Properties.Name -notContains 'Quantity') {
                    throw "Property 'Remove.File.OlderThan.Quantity' not found. Use value number '0' to move all files."
                }

                try {
                    $null = [int]$fileToRemove.OlderThan.Quantity
                }
                catch {
                    throw "Property 'Remove.File.OlderThan.Quantity' needs to be a number, the value '$($fileToRemove.OlderThan.Quantity)' is not supported. Use value number '0' to move all files."
                }
                #endregion

                if (
                    ($fileToRemove.Path -notMatch '^\\\\') -and
                    (-not $fileToRemove.ComputerName)
                ) {
                    throw "No 'Remove.File.ComputerName' found for path '$($fileToRemove.Path)'"
                }
            }

            foreach ($fileInFolderToRemove in $file.Remove.FilesInFolder) {
                @(
                    'Path', 'OlderThan'
                ).where(
                    { -not $fileInFolderToRemove.$_ }
                ).foreach(
                    { throw "Property 'Remove.FilesInFolder.$_' not found" }
                )

                #region OlderThan
                if (-not $fileInFolderToRemove.OlderThan.Unit) {
                    throw "No 'Remove.FilesInFolder.OlderThan.Unit' found"
                }

                if ($fileInFolderToRemove.OlderThan.Unit -notMatch '^Day$|^Month$|^Year$') {
                    throw "Value '$($fileInFolderToRemove.OlderThan.Unit)' is not supported by 'Remove.FilesInFolder.OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'."
                }

                if ($fileInFolderToRemove.OlderThan.PSObject.Properties.Name -notContains 'Quantity') {
                    throw "Property 'Remove.FilesInFolder.OlderThan.Quantity' not found. Use value number '0' to move all files."
                }

                try {
                    $null = [int]$fileInFolderToRemove.OlderThan.Quantity
                }
                catch {
                    throw "Property 'Remove.FilesInFolder.OlderThan.Quantity' needs to be a number, the value '$($fileInFolderToRemove.OlderThan.Quantity)' is not supported. Use value number '0' to move all files."
                }
                #endregion

                if (
                    ($fileInFolderToRemove.Path -notMatch '^\\\\') -and
                    (-not $fileInFolderToRemove.ComputerName)
                ) {
                    throw "No 'Remove.FilesInFolder.ComputerName' found for path '$($fileInFolderToRemove.Path)'"
                }

                #region Test boolean values
                foreach (
                    $boolean in
                    @(
                        'Recurse'
                    )
                ) {
                    try {
                        $null = [Boolean]::Parse($fileInFolderToRemove.$boolean)
                    }
                    catch {
                        throw "Property 'Remove.FilesInFolder.$boolean' is not a boolean value"
                    }
                }
                #endregion
            }

            foreach ($emptyFoldersToRemove in $file.Remove.EmptyFolders) {
                @(
                    'Path'
                ).where(
                    { -not $emptyFoldersToRemove.$_ }
                ).foreach(
                    { throw "Property 'Remove.EmptyFolders.$_' not found" }
                )

                if (
                    ($emptyFoldersToRemove.Path -notMatch '^\\\\') -and
                    (-not $emptyFoldersToRemove.ComputerName)
                ) {
                    throw "No 'Remove.EmptyFolders.ComputerName' found for path '$($emptyFoldersToRemove.Path)'"
                }
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        #region Convert .json file
        @(
            $file.Remove.File,
            $file.Remove.FilesInFolder,
            $file.Remove.EmptyFolders
        ).foreach(
            {
                $_.Path = $_.Path.ToLower()

                #region Set ComputerName
                if (
                    (-not $_.ComputerName) -or
                    ($_.ComputerName -eq 'localhost') -or
                    ($_.ComputerName -eq "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
                ) {
                    $_.ComputerName = $env:COMPUTERNAME
                }
                #endregion
            }
        )
        #endregion

        #region Create tasks to execute
        $tasksToExecute = @()

        $file.Remove.File.foreach(
            {
                $tasksToExecute += $_ | Select-Object -Property *,
                @{
                    Name       = 'Type'
                    Expression = { 'RemoveFile' }
                }
            }
        )

        $file.Remove.FilesInFolder.foreach(
            {
                $tasksToExecute += $_ | Select-Object -Property *,
                @{
                    Name       = 'Type'
                    Expression = { 'RemoveFilesInFolder' }
                }
            }
        )

        $file.Remove.EmptyFolders.foreach(
            {
                $tasksToExecute += $_ | Select-Object -Property *,
                @{
                    Name       = 'Type'
                    Expression = { 'RemoveEmptyFolders' }
                }
            }
        )
        #endregion
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
        $scriptBlock = {
            try {
                $task = $_

                #region Declare variables for parallel execution
                if (-not $MaxConcurrentJobs) {
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                    $EventVerboseParams = $using:EventVerboseParams
                    $EventErrorParams = $using:EventErrorParams
                    $VerbosePreference = $using:VerbosePreference
                }
                #endregion

                $invokeParams = @{
                    ArgumentList = $task.Remove, $task.Path, $task.OlderThanDays, $task.RemoveEmptyFolders
                    ScriptBlock  = $null
                }

                $M = "Start job on '{0}' with Remove '{1}' Path '{2}' OlderThanDays '{3}' RemoveEmptyFolders '{4}'" -f
                $task.ComputerName,
                $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
                $invokeParams.ArgumentList[2], $invokeParams.ArgumentList[3]
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

                #region Start job
                $computerName = $task.ComputerName

                $task.Job.Results += if (
                    $computerName -eq $ENV:COMPUTERNAME
                ) {
                    $params = $invokeParams.ArgumentList
                    & $invokeParams.ScriptBlock @params
                }
                else {
                    $invokeParams += @{
                        ConfigurationName = $PSSessionConfiguration
                        ComputerName      = $computerName
                        ErrorAction       = 'Stop'
                    }
                    Invoke-Command @invokeParams
                }
                #endregion
            }
            catch {
                $task.Job.Errors += $_
                $Error.RemoveAt(0)

                $M = "'{0}' Error for Remove '{1}' Path '{2}' OlderThanDays '{3}' RemoveEmptyFolders '{4}' Name '{5}': {6}" -f
                $task.ComputerName, $task.Remove, $task.Path,
                $task.OlderThanDays, $task.RemoveEmptyFolders,
                $task.Name, $task.Job.Errors[0]
                Write-Warning $M; Write-EventLog @EventErrorParams -Message $M
            }
        }

        #region Run code serial or parallel
        $foreachParams = if ($MaxConcurrentJobs -eq 1) {
            @{
                Process = $scriptBlock
            }
        }
        else {
            @{
                Parallel      = $scriptBlock
                ThrottleLimit = $MaxConcurrentJobs
            }
        }
        #endregion

        $Tasks | ForEach-Object @foreachParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    try {
        $mailParams = @{ }

        $excelParams = @{
            Path               = $logFile + ' - Log.xlsx'
            NoNumberConversion = '*'
            AutoSize           = $true
            FreezeTopRow       = $true
        }
        $excelSheet = @{
            Overview = @()
            Errors   = @()
        }

        #region Create Excel worksheet Overview
        $excelSheet.Overview += foreach (
            $task in
            $Tasks
        ) {
            $task.Job.Results | Select-Object -Property 'ComputerName',
            'Type',
            @{
                Name       = 'Path';
                Expression = { $_.FullName }
            },
            'CreationTime', @{
                Name       = 'OlderThanDays';
                Expression = { $task.OlderThanDays }
            }, 'Action', 'Error'
        }

        if ($excelSheet.Overview) {
            $M = "Export $($excelSheet.Overview.Count) rows to Excel"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $excelParams.WorksheetName = 'Overview'
            $excelParams.TableName = 'Overview'
            $excelSheet.Overview | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Create Excel worksheet Errors
        $excelSheet.Errors += foreach (
            $task in
            $Tasks
        ) {
            $task.Job.Errors | Where-Object { $_ } | Select-Object -Property @{
                Name       = 'ComputerName';
                Expression = { $task.ComputerName }
            },
            @{
                Name       = 'Path';
                Expression = { $task.Path }
            },
            @{
                Name       = 'Remove';
                Expression = { $task.Remove }
            },
            @{
                Name       = 'OlderThanDays';
                Expression = { $task.OlderThanDays }
            },
            @{
                Name       = 'RemoveEmptyFolders';
                Expression = { $task.RemoveEmptyFolders }
            },
            @{
                Name       = 'Error';
                Expression = { $_ -join ', ' }
            }
        }

        if ($excelSheet.Errors) {
            $excelParams.WorksheetName = 'Errors'
            $excelParams.TableName = 'Errors'

            $M = "Export {0} rows to sheet '{1}' in Excel file '{2}'" -f
            $excelSheet.Errors.Count,
            $excelParams.WorksheetName, $excelParams.Path
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $excelSheet.Errors | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Send mail to user

        #region Error counters
        $counter = @{
            removedItems  = (
                $Tasks.Job.Results |
                Where-Object { ($_.Action -eq 'Removed') } |
                Measure-Object
            ).Count
            removalErrors = (
                $Tasks.Job.Results.Error | Measure-Object
            ).Count
            jobErrors     = (
                $Tasks.Job.Errors | Measure-Object
            ).Count
            systemErrors  = (
                $Error.Exception.Message | Measure-Object
            ).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'
        $mailParams.Subject = '{0} removed' -f $counter.removedItems

        if (
            $totalErrorCount = $counter.removalErrors + $counter.jobErrors +
            $counter.systemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion

        #region Create html lists
        $systemErrorsHtmlList = if ($counter.systemErrors) {
            "<p>Detected <b>{0} non terminating error{1}</b>:{2}</p>" -f $counter.systemErrors,
            $(
                if ($counter.systemErrors -ne 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } |
                ConvertTo-HtmlListHC
            )
        }

        $jobResultsHtmlListItems = foreach (
            $task in
            $Tasks | Sort-Object -Property 'Name', 'Path', 'ComputerName'
        ) {
            "{0}<br>{1}<br>Removed: {2}{3}" -f
            $(
                if ($task.Path -match '^\\\\') {
                    '<a href="{0}">{1}</a>' -f $task.Path, $(
                        if ($task.Name) { $task.Name }
                        else { $task.Path }
                    )
                }
                else {
                    $uncPath = $task.Path -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $task.ComputerName, $task.Path[0]
                    )
                    '<a href="{0}">{1}</a>' -f $uncPath, $(
                        if ($task.Name) { $task.Name }
                        else { $uncPath }
                    )
                }
            ),
            $(
                $description = if ($task.Remove -eq 'File') {
                    if ($task.OlderThanDays -eq 0) {
                        'Remove file'
                    }
                    else {
                        "Remove file when it's older than {0} days" -f
                        $task.OlderThanDays
                    }
                }
                elseif ($task.Remove -eq 'Folder') {
                    if ($task.OlderThanDays -eq 0) {
                        'Remove folder'
                    }
                    else {
                        "Remove folder when it's older than {0} days" -f
                        $task.OlderThanDays
                    }
                }
                elseif ($task.Remove -eq 'Content') {
                    if ($task.OlderThanDays -eq 0) {
                        'Remove folder content'
                    }
                    else {
                        'Remove folder content that is older than {0} days' -f
                        $task.OlderThanDays
                    }
                }
                if ($task.RemoveEmptyFolders) {
                    $description += ' and remove empty folders'
                }
                $description
            ),
            $(
                (
                    $task.Job.Results |
                    Where-Object { $_.Action -eq 'Removed' } |
                    Measure-Object
                ).Count
            ),
            $(
                if ($errorCount = (
                        $task.Job.Results | Where-Object { $_.Error } |
                        Measure-Object
                    ).Count + $task.Job.Errors.Count) {
                    ', <b style="color:red;">errors: {0}</b>' -f $errorCount
                }
            )
        }

        $jobResultsHtmlList = $jobResultsHtmlListItems |
        ConvertTo-HtmlListHC -Spacing Wide
        #endregion

        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $systemErrorsHtmlList
                <p>Summary:</p>
                $jobResultsHtmlList"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        if ($mailParams.Attachments) {
            $mailParams.Message +=
            "<p><i>* Check the attachment for details</i></p>"
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}