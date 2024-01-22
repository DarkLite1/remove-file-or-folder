#Requires -Version 5.1
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.Remoting

<#
.SYNOPSIS
    Remove files, folders or folder content on remote machines.

.DESCRIPTION
    This script reads a .JSON file containing the destination paths where files
    or folders need to be removed. When 'OlderThanDays' is '0' all files or
    folders will be removed, depending on the chosen 'Remove' type, regardless
    their creation date.

.PARAMETER MailTo
    E-mail addresses of where to send the summary e-mail

.PARAMETER Destinations
    Contains an array of objects where each object represents a 'Path' and its
    specific settings on file, folder or content removal

.PARAMETER Destinations.Name
    The name to display in the email send to the user instead of the full path

.PARAMETER Destinations.Remove
    file    : remove the file specified in 'Path'
    folder  : remove the files in the 'Path' folder
    content : remove the files in the 'Path' folder and sub folders
    (use RemoveEmptyFolders to delete empty folder)

.PARAMETER Destinations.Path
    Can be a local path when 'ComputerName' is used or a UNC path

.PARAMETER Destinations.OlderThanDays
    Only remove files that are older than x days

.PARAMETER Destinations.RemoveEmptyFolders
    Can only be used with 'Remove' set to 'content' and will remove all empty
    folders after the files have been removed
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Remove file or folder\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
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

        $compareDate = (Get-Date).AddDays(-$OlderThanDays)

        #region Create get params and test file folder
        $commandToRun = "Get-Item -LiteralPath '$Path'"
        $removalType = 'File'

        switch ($Type) {
            'file' {
                if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
                    [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        Type         = 'File'
                        FullName     = $Path
                        CreationTime = $null
                        Action       = $null
                        Error        = 'Path not found'
                    }
                    Exit
                }
            }
            'folder' {
                if (-not (Test-Path -LiteralPath $Path -PathType Container)
                ) {
                    [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        Type         = 'Folder'
                        FullName     = $Path
                        CreationTime = $null
                        Action       = $null
                        Error        = 'Path not found'
                    }
                    Exit
                }
                $removalType = 'Folder'
                break
            }
            'content' {
                if (
                    -not (Test-Path -LiteralPath $Path -PathType Container)
                ) {
                    throw "Folder '$Path' not found"
                }
                $commandToRun = "Get-ChildItem -LiteralPath '$Path' -Recurse -File"
                break
            }
            Default {
                throw "Type '$_' not supported"
            }
        }
        #endregion

        #region Remove items
        $removeParams = @{
            Recurse     = $true
            Force       = $true
            ErrorAction = 'Stop'
        }

        Invoke-Expression $commandToRun | Where-Object {
            ($_.CreationTime -lt $compareDate) -or ($OlderThanDays -eq 0)
        } | ForEach-Object {
            try {
                $result = [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    Type         = $removalType
                    FullName     = $_.FullName
                    CreationTime = $_.CreationTime
                    Action       = $null
                    Error        = $null
                }

                Remove-Item @removeParams -LiteralPath $_.FullName
                $result.Action = 'Removed'
            }
            catch {
                $result.Error = $_
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
        #endregion

        #region Remove empty folders
        if (($Type -eq 'content') -and ($RemoveEmptyFolders)) {
            $failedFolderRemoval = @()

            $getParams = @{
                LiteralPath = $Path
                Directory   = $true
                Recurse     = $true
            }

            while (
                $emptyFolders = Get-ChildItem @getParams |
                Where-Object {
                    ($_.GetFileSystemInfos().Count -eq 0) -and
                    ($failedFolderRemoval -notContains $_.FullName)
                }
            ) {
                $emptyFolders | ForEach-Object {
                    try {
                        $result = [PSCustomObject]@{
                            ComputerName = $env:COMPUTERNAME
                            Type         = 'Folder'
                            FullName     = $_.FullName
                            CreationTime = $_.CreationTime
                            Action       = $null
                            Error        = $null
                        }

                        Remove-Item @removeParams -LiteralPath $_.FullName
                        $result. Action = 'Removed'
                    }
                    catch {
                        $result.Error = $_
                        $Error.RemoveAt(0)
                        $failedFolderRemoval += $_.FullName
                    }
                    finally {
                        $result
                    }
                }
            }
        }
        #endregion
    }

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
        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }

        if (-not ($MaxConcurrentJobs = $file.MaxConcurrentJobs)) {
            throw "Property 'MaxConcurrentJobs' not found"
        }
        try {
            $null = $MaxConcurrentJobs.ToInt16($null)
        }
        catch {
            throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$MaxConcurrentJobs' is not supported."
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
                throw "Input file '$ImportFile' destination path '$($d.Path)': No 'ComputerName' found."
            }
            #endregion

            #region OlderThanDays
            if ($d.PSObject.Properties.Name -notContains 'OlderThanDays') {
                throw "Input file '$ImportFile' destination path '$($d.Path)': Property 'OlderThanDays' not found. Use value number '0' to remove all."
            }
            try {
                $null = $d.OlderThanDays.ToInt16($null)
            }
            catch {
                throw "Input file '$ImportFile' destination path '$($d.Path)': Property 'OlderThanDays' needs to be a number, the value '$($d.OlderThanDays)' is not supported. Use value number '0' to remove all."
            }
            #endregion

            #region Remove
            if ($d.PSObject.Properties.Name -notContains 'Remove') {
                throw "Input file '$ImportFile' destination path '$($d.Path)': Property 'Remove' not found. Valid values are 'folder', 'file' or 'content'."
            }
            if ($d.Remove -notMatch '^folder$|^file$|^content$') {
                throw "Input file '$ImportFile' destination path '$($d.Path)': Value '$($d.Remove)' in 'Remove' is not valid, only values 'folder', 'file' or 'content' are supported."
            }
            #endregion

            #region RemoveEmptyFolders
            if ($d.Remove -eq 'content') {
                if (
                    $d.PSObject.Properties.Name -notContains 'RemoveEmptyFolders'
                ) {
                    throw "Input file '$ImportFile' destination path '$($d.Path)': Property 'RemoveEmptyFolders' not found."
                }
                if (-not ($d.RemoveEmptyFolders -is [boolean])) {
                    throw "Input file '$ImportFile' destination path '$($d.Path)': The value '$($d.RemoveEmptyFolders)' in 'RemoveEmptyFolders' is not a true false value."
                }
            }
            else {
                if ($d.RemoveEmptyFolders) {
                    throw "Input file '$ImportFile' destination path '$($d.Path)': Property 'RemoveEmptyFolders' cannot be used with 'Remove' value '$($d.Remove)'."
                }
            }
            #endregion
        }
        #endregion

        #region Convert .json properties
        foreach ($d in $Destinations) {
            $d.Remove = $d.Remove.ToLower()
            $d.Path = $d.Path.ToLower()
            if ($d.ComputerName) {
                $d.ComputerName = $d.ComputerName.ToUpper()
            }
            if ($d.Remove -ne 'content') {
                $addParams = @{
                    InputObject       = $d
                    NotePropertyName  = 'RemoveEmptyFolders'
                    NotePropertyValue = $false
                }
                Add-Member @addParams
            }
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
        foreach ($d in $Destinations) {
            $invokeParams = @{
                ScriptBlock  = $scriptBlock
                ArgumentList = $d.Remove, $d.Path, $d.OlderThanDays, $d.RemoveEmptyFolders
            }

            $M = "Start job on '{0}' with Remove '{1}' Path '{2}' OlderThanDays '{3}' RemoveEmptyFolders '{4}'" -f $(
                if ($d.ComputerName) { $d.ComputerName }
                else { $env:COMPUTERNAME }
            ),
            $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
            $invokeParams.ArgumentList[2], $invokeParams.ArgumentList[3]
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            #region Add job properties
            $addParams = @{
                NotePropertyMembers = @{
                    Job        = $null
                    JobResults = @()
                    JobErrors  = @()
                }
            }
            $d | Add-Member @addParams
            #endregion

            $d.Job = if ($d.ComputerName) {
                $invokeParams.ComputerName = $d.ComputerName
                $invokeParams.AsJob = $true
                Invoke-Command @invokeParams
            }
            else {
                Start-Job @invokeParams
            }
            # & $scriptBlock -Type $d.Remove -Path $d.Path -OlderThanDays $d.OlderThanDays -RemoveEmptyFolders $d.RemoveEmptyFolders

            $waitParams = @{
                Name       = $Destinations.Job | Where-Object { $_ }
                MaxThreads = $MaxConcurrentJobs
            }
            Wait-MaxRunningJobsHC @waitParams
        }
        #endregion

        $M = "Wait for all $($Destinations.count) jobs to finish"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $null = $Destinations.Job | Wait-Job

        foreach ($d in $Destinations) {
            #region Get job results and job errors
            $jobErrors = @()
            $receiveParams = @{
                ErrorVariable = 'jobErrors'
                ErrorAction   = 'SilentlyContinue'
            }
            $d.JobResults += $d.Job | Receive-Job @receiveParams

            foreach ($e in $jobErrors) {
                $d.JobErrors += $e.ToString()
                $Error.Remove($e)

                $M = "'{0}' Error for Remove '{1}' Path '{2}' OlderThanDays '{3}' RemoveEmptyFolders '{4}' Name '{5}': {6}" -f
                $d.Job.Location, $d.Remove, $d.Path, $d.OlderThanDays,
                $d.RemoveEmptyFolders, $d.Name, $e.ToString()
                Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
            }
            #endregion
        }

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
        $excelSheet.Overview += foreach ($d in $Destinations) {
            $d.JobResults | Select-Object -Property 'ComputerName',
            'Type',
            @{
                Name       = 'Path';
                Expression = { $_.FullName }
            },
            'CreationTime', @{
                Name       = 'OlderThanDays';
                Expression = { $d.OlderThanDays }
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
        $excelSheet.Errors += foreach ($d in $Destinations) {
            $d.JobErrors | Where-Object { $_ } | Select-Object -Property @{
                Name       = 'ComputerName';
                Expression = { $d.ComputerName }
            },
            @{
                Name       = 'Path';
                Expression = { $d.Path }
            },
            @{
                Name       = 'Remove';
                Expression = { $d.Remove }
            },
            @{
                Name       = 'OlderThanDays';
                Expression = { $d.OlderThanDays }
            },
            @{
                Name       = 'RemoveEmptyFolders';
                Expression = { $d.RemoveEmptyFolders }
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
        #region Send mail to user

        #region Error counters
        $counter = @{
            removedItems  = (
                $Destinations.JobResults |
                Where-Object { ($_.Action -eq 'Removed') } |
                Measure-Object
            ).Count
            removalErrors = (
                $Destinations.JobResults.Error | Measure-Object
            ).Count
            jobErrors     = (
                $Destinations.JobErrors | Measure-Object
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
            $d in
            $Destinations | Sort-Object -Property 'Name', 'Path', 'ComputerName'
        ) {
            "{0}<br>{1}<br>Removed: {2}{3}" -f
            $(
                if ($d.Path -match '^\\\\') {
                    '<a href="{0}">{1}</a>' -f $d.Path, $(
                        if ($d.Name) { $d.Name }
                        else { $d.Path }
                    )
                }
                else {
                    $uncPath = $d.Path -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $d.ComputerName, $d.Path[0]
                    )
                    '<a href="{0}">{1}</a>' -f $uncPath, $(
                        if ($d.Name) { $d.Name }
                        else { $uncPath }
                    )
                }
            ),
            $(
                $description = if ($d.Remove -eq 'File') {
                    if ($d.OlderThanDays -eq 0) {
                        'Remove file'
                    }
                    else {
                        "Remove file when it's older than {0} days" -f
                        $d.OlderThanDays
                    }
                }
                elseif ($d.Remove -eq 'Folder') {
                    if ($d.OlderThanDays -eq 0) {
                        'Remove folder'
                    }
                    else {
                        "Remove folder when it's older than {0} days" -f
                        $d.OlderThanDays
                    }
                }
                elseif ($d.Remove -eq 'Content') {
                    if ($d.OlderThanDays -eq 0) {
                        'Remove folder content'
                    }
                    else {
                        'Remove folder content that is older than {0} days' -f
                        $d.OlderThanDays
                    }
                }
                if ($d.RemoveEmptyFolders) {
                    $description += ' and remove empty folders'
                }
                $description
            ),
            $(
                (
                    $d.JobResults |
                    Where-Object { $_.Action -eq 'Removed' } |
                    Measure-Object
                ).Count
            ),
            $(
                if ($errorCount = (
                        $d.JobResults | Where-Object { $_.Error } |
                        Measure-Object
                    ).Count + $d.JobErrors.Count) {
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