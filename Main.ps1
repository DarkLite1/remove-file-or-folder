#Requires -Version 7
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Remove selected files and empty subfolders using a JSON configuration.

.DESCRIPTION
    Reads the 'Tasks' array in a JSON file. Each task selects either specific
    'Files' or the files and empty subfolders below 'Folders'. Tasks can run
    on this computer, on remote computers, or against UNC shares.

    'OlderThan.BasedOn' selects CreationTime or LastWriteTime. Day, Month and
    Year use calendar cutoffs, not elapsed time. Quantity 0 disables the age
    filter; exclusions and the task's Recurse setting still apply.

    Empty-folder cleanup runs after all file-removal jobs finish. It checks
    subfolders at every depth, preserves each root folder, and only deletes
    folders that are empty. Age settings do not apply to folders.

    'MaxConcurrent' limits total jobs and jobs per target computer. WinRM
    abort error 995, or its matching message, allows up to three attempts
    with five seconds between attempts. Other errors are not retried.

    'Settings' controls optional email, log files and Windows event logging.
    See README.md for setup and Example.json's '?' section for each setting.
    This script deletes items; it has no preview or WhatIf mode.

    Exits with code 1 for job or script errors. On success it returns normally.

.PARAMETER ConfigurationJsonFile
    Path to the JSON configuration file. Relative paths are resolved from
    the current working directory.

.PARAMETER RemoveItemsScript
    Path to the worker script. Defaults to 'Remove items.ps1' beside this
    script. A replacement must accept the same positional parameters and
    return the same removal/error objects; it must also work independently
    when sent to a remote computer with Invoke-Command -FilePath.

.EXAMPLE
    & '.\Main.ps1' -ConfigurationJsonFile '.\MyConfig.json'

    Runs the removal tasks and reporting configured in MyConfig.json.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ConfigurationJsonFile,
    [String]$RemoveItemsScript = "$PSScriptRoot\Remove items.ps1"
)

Begin {
    . (Join-Path -Path $PSScriptRoot -ChildPath 'Functions.ps1')

    $eventLogData = [System.Collections.Generic.List[PSObject]]::new()
    $systemErrors = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date

    Try {
        $eventLogData.Add(
            [PSCustomObject]@{
                Message   = 'Script started'
                DateTime  = $scriptStartTime
                EntryType = 'Information'
                EventID   = '100'
            }
        )

        #region Import .json file
        Write-Verbose "Import .json file '$ConfigurationJsonFile'"

        $jsonFileItem = Get-Item -LiteralPath $ConfigurationJsonFile -ErrorAction Stop

        $file = Get-Content -LiteralPath $jsonFileItem -Raw -Encoding UTF8 -ErrorAction Stop |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        try {
            #region Test Settings
            $settings = $file.Settings

            if (-not $settings.ScriptName) {
                throw "Property 'Settings.ScriptName' not found"
            }

            #region SendMail
            $sendMail = $settings.SendMail

            if (-not $sendMail.When) {
                throw "Property 'Settings.SendMail.When' not found"
            }

            if ($sendMail.When -notin 'Never', 'Always', 'OnError', 'OnErrorOrAction') {
                throw "Property 'Settings.SendMail.When' with value '$($sendMail.When)' is not supported. Supported values are 'Never', 'Always', 'OnError' or 'OnErrorOrAction'."
            }

            if ($sendMail.When -ne 'Never') {
                $mandatoryMailProperties = [ordered]@{
                    'From'                 = $sendMail.From
                    'Smtp.ServerName'      = $sendMail.Smtp.ServerName
                    'Smtp.Port'            = $sendMail.Smtp.Port
                    'AssemblyPath.MailKit' = $sendMail.AssemblyPath.MailKit
                    'AssemblyPath.MimeKit' = $sendMail.AssemblyPath.MimeKit
                }

                foreach ($property in $mandatoryMailProperties.GetEnumerator()) {
                    if (-not $property.Value) {
                        throw "Property 'Settings.SendMail.$($property.Key)' not found"
                    }
                }

                if (-not ($sendMail.To -or $sendMail.Bcc)) {
                    throw "Property 'Settings.SendMail.To' or 'Settings.SendMail.Bcc' not found"
                }

                if (
                    ($sendMail.Smtp.Port -notmatch '^ENV:') -and
                    ($sendMail.Smtp.Port -notin 25, 465, 587, 2525)
                ) {
                    throw "Property 'Settings.SendMail.Smtp.Port' with value '$($sendMail.Smtp.Port)' is not supported. Supported values are 25, 465, 587 or 2525."
                }

                if (
                    $sendMail.Smtp.ConnectionType -and
                    ($sendMail.Smtp.ConnectionType -notmatch '^ENV:') -and
                    ($sendMail.Smtp.ConnectionType -notin 'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')
                ) {
                    throw "Property 'Settings.SendMail.Smtp.ConnectionType' with value '$($sendMail.Smtp.ConnectionType)' is not supported. Supported values are 'None', 'Auto', 'SslOnConnect', 'StartTls' or 'StartTlsWhenAvailable'."
                }
            }
            #endregion

            #region SaveLogFiles
            $saveLogFiles = $settings.SaveLogFiles

            if (
                ($null -ne $saveLogFiles.DeleteLogsAfterDays) -and
                ("$($saveLogFiles.DeleteLogsAfterDays)" -notmatch '^\d+$')
            ) {
                throw "Property 'Settings.SaveLogFiles.DeleteLogsAfterDays' needs to be a positive number, the value '$($saveLogFiles.DeleteLogsAfterDays)' is not supported."
            }
            #endregion

            #region SaveInEventLog
            $saveInEventLog = $settings.SaveInEventLog

            if ($null -eq $saveInEventLog.Save) {
                throw "Property 'Settings.SaveInEventLog.Save' not found"
            }
            if ($saveInEventLog.Save -isnot [bool]) {
                throw "Property 'Settings.SaveInEventLog.Save' needs to be true or false, the value '$($saveInEventLog.Save)' is not supported."
            }
            if ($saveInEventLog.Save -and (-not $saveInEventLog.LogName)) {
                throw "Property 'Settings.SaveInEventLog.LogName' not found"
            }
            #endregion
            #endregion

            if ($file.PSObject.Properties.Name -contains 'Remove') {
                throw "Property 'Remove' is no longer supported, use 'Tasks' instead. See 'Example.json'."
            }

            @(
                'MaxConcurrent', 'Tasks'
            ).where(
                { -not $file.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            foreach ($property in 'JobsTotal', 'JobsPerComputer') {
                $value = $file.MaxConcurrent.$property

                if ($null -eq $value) {
                    throw "Property 'MaxConcurrent.$property' not found"
                }

                if (
                    ("$value" -notMatch '^\d+$') -or ([int]"$value" -lt 1)
                ) {
                    throw "Property 'MaxConcurrent.$property' needs to be a number of 1 or higher, the value '$value' is not supported."
                }
            }

            $maxConcurrentJobsTotal = [int]$file.MaxConcurrent.JobsTotal
            $maxConcurrentJobsPerComputer = [int]$file.MaxConcurrent.JobsPerComputer

            #region Tasks
            $tasksToExecute = @()
            $tasks = @($file.Tasks)

            for ($i = 0; $i -lt $tasks.Count; $i++) {
                $task = $tasks[$i]
                $prefix = "Tasks[$i]"
                $taskProperties = $task.PSObject.Properties.Name

                foreach ($name in $taskProperties) {
                    if ($name -notin '?', 'ComputerName', 'Files', 'Folders', 'ExcludeFolders', 'ExcludeFiles', 'OlderThan', 'Recurse', 'RemoveEmptyFolders') {
                        throw "Property '$prefix.$name' is not supported"
                    }
                }

                #region Files or Folders
                if ($task.Files -and $task.Folders) {
                    throw "Property '$prefix.Files' and '$prefix.Folders' cannot be used at the same time"
                }
                if (-not ($task.Files -or $task.Folders)) {
                    throw "Property '$prefix.Files' or '$prefix.Folders' not found"
                }

                $listName = if ($task.Files) { 'Files' } else { 'Folders' }
                $removeFiles = $taskProperties -contains 'OlderThan'
                #endregion

                #region Recurse and RemoveEmptyFolders
                if ($listName -eq 'Files') {
                    if (-not $removeFiles) {
                        throw "Property '$prefix.OlderThan' not found"
                    }

                    foreach ($name in 'Recurse', 'RemoveEmptyFolders') {
                        if ($taskProperties -contains $name) {
                            throw "Property '$prefix.$name' can only be used with '$prefix.Folders'"
                        }
                    }
                }
                else {
                    $booleans = @('RemoveEmptyFolders')

                    if ($removeFiles) {
                        $booleans += 'Recurse'
                    }
                    elseif ($taskProperties -contains 'Recurse') {
                        throw "Property '$prefix.Recurse' can only be used together with '$prefix.OlderThan'"
                    }

                    foreach ($name in $booleans) {
                        if ($null -eq $task.$name) {
                            throw "Property '$prefix.$name' not found"
                        }
                        if ($task.$name -isnot [bool]) {
                            throw "Property '$prefix.$name' needs to be true or false, the value '$($task.$name)' is not supported."
                        }
                    }

                    if ((-not $removeFiles) -and (-not $task.RemoveEmptyFolders)) {
                        throw "Property '$prefix.OlderThan' not found. Use 'OlderThan' to remove files, 'RemoveEmptyFolders' to remove empty folders or both."
                    }
                }
                #endregion

                #region OlderThan
                if ($removeFiles) {
                    if (-not $task.OlderThan.Unit) {
                        throw "Property '$prefix.OlderThan.Unit' not found"
                    }

                    if ($task.OlderThan.Unit -notin 'Day', 'Month', 'Year') {
                        throw "Property '$prefix.OlderThan.Unit' with value '$($task.OlderThan.Unit)' is not supported. Supported values are 'Day', 'Month' or 'Year'."
                    }

                    if ($task.OlderThan.PSObject.Properties.Name -notContains 'Quantity') {
                        throw "Property '$prefix.OlderThan.Quantity' not found. Use value 0 to remove all files."
                    }

                    if ("$($task.OlderThan.Quantity)" -notMatch '^\d+$') {
                        throw "Property '$prefix.OlderThan.Quantity' needs to be a positive number, the value '$($task.OlderThan.Quantity)' is not supported. Use value 0 to remove all files."
                    }

                    if (-not $task.OlderThan.BasedOn) {
                        throw "Property '$prefix.OlderThan.BasedOn' not found. Use 'CreationTime' or 'LastWriteTime'."
                    }

                    if ($task.OlderThan.BasedOn -notin 'CreationTime', 'LastWriteTime') {
                        throw "Property '$prefix.OlderThan.BasedOn' with value '$($task.OlderThan.BasedOn)' is not supported. Supported values are 'CreationTime' or 'LastWriteTime'."
                    }

                    foreach ($name in $task.OlderThan.PSObject.Properties.Name) {
                        if ($name -notin 'Quantity', 'Unit', 'BasedOn') {
                            throw "Property '$prefix.OlderThan.$name' is not supported"
                        }
                    }
                }
                #endregion

                $computerName = if (
                    (-not $task.ComputerName) -or
                    ($task.ComputerName -eq 'localhost') -or
                    ($task.ComputerName -eq "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
                ) {
                    $env:COMPUTERNAME
                }
                else {
                    $task.ComputerName
                }

                #region Create one task to execute per path
                $entries = @($task.$listName)
                $folderPaths = @()

                #region ExcludeFolders and ExcludeFiles
                $excludes = @{
                    ExcludeFolders = @()
                    ExcludeFiles   = @()
                }

                foreach ($excludeName in 'ExcludeFolders', 'ExcludeFiles') {
                    if ($taskProperties -notcontains $excludeName) { continue }

                    if ($listName -eq 'Files') {
                        throw "Property '$prefix.$excludeName' can only be used with '$prefix.Folders'"
                    }
                    if (($excludeName -eq 'ExcludeFiles') -and (-not $removeFiles)) {
                        throw "Property '$prefix.ExcludeFiles' can only be used together with '$prefix.OlderThan'"
                    }

                    $excludes[$excludeName] = @($task.$excludeName)

                    foreach ($excludePath in $excludes[$excludeName]) {
                        if (($excludePath -isnot [string]) -or (-not $excludePath)) {
                            $kind = if ($excludeName -eq 'ExcludeFiles') { 'file' } else { 'folder' }
                            throw "Property '$prefix.$excludeName' needs to be an array of $kind paths, the value '$excludePath' is not supported."
                        }
                    }
                    $excludes[$excludeName] = @(
                        foreach ($excludePath in $excludes[$excludeName]) {
                            $excludePath = $excludePath.Replace('/', '\')
                            if ([System.IO.Path]::IsPathFullyQualified($excludePath)) {
                                [System.IO.Path]::GetFullPath($excludePath)
                            }
                            else { $excludePath }
                        }
                    )
                }
                #endregion

                for ($j = 0; $j -lt $entries.Count; $j++) {
                    $entry = $entries[$j]
                    $entryPrefix = "$prefix.$listName[$j]"

                    if ($entry -is [string]) {
                        $name = $null
                        $path = $entry
                    }
                    else {
                        foreach ($property in $entry.PSObject.Properties.Name) {
                            if ($property -notin 'Name', 'Path') {
                                throw "Property '$entryPrefix.$property' is not supported"
                            }
                        }

                        $name = $entry.Name
                        $path = $entry.Path
                    }

                    if (-not $path) {
                        throw "Property '$entryPrefix' needs a path"
                    }

                    $path = $path.Replace('/', '\')
                    if ([System.IO.Path]::IsPathFullyQualified($path)) {
                        $path = [System.IO.Path]::GetFullPath($path)
                    }

                    if (($path -notMatch '^\\\\') -and (-not $task.ComputerName)) {
                        throw "Property '$prefix.ComputerName' not found, it is required for the local path '$path'"
                    }

                    $types = if ($listName -eq 'Files') {
                        'RemoveFile'
                    }
                    else {
                        if ($removeFiles) { 'RemoveFilesInFolder' }
                        if ($task.RemoveEmptyFolders) { 'RemoveEmptyFolders' }
                    }

                    $folderPaths += $path.TrimEnd('\')
                    $pathPrefix = "$($path.TrimEnd('\'))\"
                    $pathExcludes = @{}

                    foreach ($excludeName in 'ExcludeFolders', 'ExcludeFiles') {
                        $pathExcludes[$excludeName] = @(
                            $excludes[$excludeName].Where({
                                    $_.TrimEnd('\').StartsWith($pathPrefix, [StringComparison]::OrdinalIgnoreCase)
                                })
                        )
                    }

                    foreach ($type in $types) {
                        $tasksToExecute += [PSCustomObject]@{
                            Name           = $name
                            ComputerName   = $computerName
                            Path           = $path
                            Type           = $type
                            OlderThan      = if ($type -ne 'RemoveEmptyFolders') { $task.OlderThan }
                            Recurse        = $task.Recurse
                            ExcludeFolders = $pathExcludes.ExcludeFolders
                            ExcludeFiles   = $pathExcludes.ExcludeFiles
                            Job            = @{
                                Results = @()
                                Errors  = @()
                            }
                        }
                    }
                }

                foreach ($excludeName in 'ExcludeFolders', 'ExcludeFiles') {
                    foreach ($excludePath in $excludes[$excludeName]) {
                        $isBelowFolder = $folderPaths.Where({
                                $excludePath.TrimEnd('\').StartsWith("$_\", [StringComparison]::OrdinalIgnoreCase)
                            })

                        if (-not $isBelowFolder) {
                            $kind = if ($excludeName -eq 'ExcludeFiles') { 'a file' } else { 'a subfolder' }
                            throw "Property '$prefix.$excludeName' contains '$excludePath', which is not $kind of a path in '$prefix.Folders'"
                        }
                    }
                }
                #endregion
            }
            #endregion
        }
        catch {
            throw "Input file '$ConfigurationJsonFile': $_"
        }
        #endregion

        #region Test path exists
        try {
            $removeItemsScriptPath = (
                Get-Item -LiteralPath $RemoveItemsScript -ErrorAction Stop
            ).FullName
        }
        catch {
            throw "RemoveItemsScript '$RemoveItemsScript' not found"
        }
        #endregion

        #region Convert .json file
        $PSSessionConfiguration = $file.PSSessionConfiguration

        if (-not $PSSessionConfiguration) {
            $PSSessionConfiguration = 'PowerShell.7'
        }
        #endregion

        if (-not $tasksToExecute) {
            throw 'No tasks to execute'
        }
        #endregion
    }
    Catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "$_"
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
}

Process {
    if ($systemErrors) { return }

    Try {
        #region Create DTOs
        $taskDtos = for ($i = 0; $i -lt $tasksToExecute.Count; $i++) {
            $task = $tasksToExecute[$i]

            switch ($task.Type) {
                'RemoveFile' {
                    $argumentList = @(
                        'File', $task.Path, @(),
                        $task.OlderThan.Unit, $task.OlderThan.Quantity,
                        $false, @(), $task.OlderThan.BasedOn
                    )

                    $M = "Prepared job '$_' on '{0}' with Path '{1}' OlderThan.Quantity '{3}' OlderThan.Unit '{2}' OlderThan.BasedOn '{4}'" -f
                    $task.ComputerName,
                    $argumentList[1], $argumentList[3], $argumentList[4],
                    $argumentList[7]

                    break
                }
                'RemoveFilesInFolder' {
                    $argumentList = @(
                        'FilesInFolder', $task.Path, $task.ExcludeFolders,
                        $task.OlderThan.Unit, $task.OlderThan.Quantity,
                        $task.Recurse, $task.ExcludeFiles, $task.OlderThan.BasedOn
                    )

                    $M = "Prepared job '$_' on '{0}' with Path '{1}' OlderThan.Quantity '{3}' OlderThan.Unit '{2}' OlderThan.BasedOn '{7}' Recurse '{4}' ExcludeFolders '{5}' ExcludeFiles '{6}'" -f
                    $task.ComputerName,
                    $argumentList[1], $argumentList[3], $argumentList[4],
                    $argumentList[5], ($task.ExcludeFolders -join "', '"),
                    ($task.ExcludeFiles -join "', '"), $argumentList[7]

                    break
                }
                'RemoveEmptyFolders' {
                    $argumentList = @(
                        'EmptyFolders', $task.Path, $task.ExcludeFolders
                    )

                    $M = "Prepared job '$_' on '{0}' with Path '{1}' ExcludeFolders '{2}'" -f
                    $task.ComputerName, $argumentList[1],
                    ($task.ExcludeFolders -join "', '")

                    break
                }
                Default {
                    throw "Type '$_' not supported"
                }
            }

            Write-Verbose $M

            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = $M
                    DateTime  = Get-Date
                    EntryType = 'Information'
                    EventID   = '4'
                }
            )

            # local jobs on a UNC path load the file server, not this computer
            $target = if (
                ($task.ComputerName -eq $env:COMPUTERNAME) -and
                ($task.Path -match '^\\\\([^\\]+)')
            ) {
                $Matches[1]
            }
            else {
                $task.ComputerName
            }

            [PSCustomObject]@{
                ID           = $i
                Type         = $task.Type
                ComputerName = $task.ComputerName
                Target       = $target
                FilePath     = $removeItemsScriptPath
                ArgumentList = $argumentList
                StartMessage = $M
            }
        }
        #endregion

        $workerScriptBlock = {
            param (
                [Parameter(Mandatory)]
                [PSCustomObject]$Worker,
                [Parameter(Mandatory)]
                [String]$SessionConfiguration,
                [Parameter(Mandatory)]
                [Int]$MaxAttempts,
                [Parameter(Mandatory)]
                [Int]$RetryDelaySeconds,
                [Parameter(Mandatory)]
                [System.Management.Automation.ActionPreference]$WorkerVerbosePreference
            )

            $VerbosePreference = $WorkerVerbosePreference
            $writeVerbose = $VerbosePreference -notin 'SilentlyContinue', 'Ignore'
            $sessionOption = New-PSSessionOption -MaximumReceivedObjectSize ([Int32]::MaxValue)

            # TryDequeue is atomic, so workers of the same computer share one queue
            $dto = $null

            while ($Worker.JobQueue.TryDequeue([ref]$dto)) {
                $result = [PSCustomObject]@{
                    ID      = $dto.ID
                    Results = @()
                    Errors  = @()
                }

                $attempt = 0

                while ($true) {
                    $attempt++
                    $session = $null
                    $needsRetry = $false

                    try {
                        if ($writeVerbose) {
                            Write-Verbose "Starting job '$($dto.Type)' on '$($dto.ComputerName)' for '$($dto.ArgumentList[1])' (attempt $attempt of $MaxAttempts)"
                        }
                        if ($dto.ComputerName -eq $env:COMPUTERNAME) {
                            $jobStage = 'Run local worker'
                            $arguments = $dto.ArgumentList
                            $result.Results = @(& $dto.FilePath @arguments)
                        }
                        else {
                            $jobStage = 'Open remote session'
                            $sessionParams = @{
                                ComputerName      = $dto.ComputerName
                                ConfigurationName = $SessionConfiguration
                                SessionOption     = $sessionOption
                                ErrorAction       = 'Stop'
                            }
                            $session = New-PSSession @sessionParams

                            $jobStage = 'Run remote worker'
                            $invokeParams = @{
                                Session      = $session
                                FilePath     = $dto.FilePath
                                ArgumentList = $dto.ArgumentList
                                ErrorAction  = 'Stop'
                                Verbose      = $writeVerbose
                            }
                            $result.Results = @(Invoke-Command @invokeParams)
                        }
                    }
                    catch {
                        # Win32 995: client-side WinRM abort, not a removal failure
                        $isTransientAbort = (
                            ("$($_.Exception.Message)" -match 'I/O operation has been aborted') -or
                            (($_.Exception.HResult -band 0xFFFF) -eq 995)
                        )

                        if ($isTransientAbort -and ($attempt -lt $MaxAttempts)) {
                            $needsRetry = $true
                        }
                        else {
                            $_ | Add-Member -NotePropertyName JobStage -NotePropertyValue $jobStage -Force
                            $result.Errors = @($_)
                        }

                        $Error.RemoveAt(0)
                    }
                    finally {
                        # closing the session stops an orphaned remote command before a retry
                        if ($session) {
                            Remove-PSSession -Session $session -ErrorAction SilentlyContinue
                        }
                    }

                    if ($needsRetry) {
                        if ($writeVerbose) {
                            Write-Verbose "Retrying job '$($dto.Type)' on '$($dto.ComputerName)' for '$($dto.ArgumentList[1])' after WinRM abort; attempt $($attempt + 1) of $MaxAttempts in $RetryDelaySeconds seconds"
                        }
                        Start-Sleep -Seconds $RetryDelaySeconds
                        continue
                    }

                    break
                }

                $result
            }
        }

        $params = @{
            ScriptBlock   = $workerScriptBlock
            ThrottleLimit = $maxConcurrentJobsTotal
            ArgumentList  = $PSSessionConfiguration, 3, 5, $VerbosePreference
        }

        # empty folders can only be removed after the files are removed
        $jobResults = foreach ($emptyFoldersPhase in $false, $true) {
            $workers = $taskDtos.Where(
                { ($_.Type -eq 'RemoveEmptyFolders') -eq $emptyFoldersPhase }
            ) | Group-Object -Property 'Target' | ForEach-Object {
                $group = $_
                $jobQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()

                foreach ($dto in $group.Group) {
                    $jobQueue.Enqueue($dto)
                }

                $workerCount = [math]::Min(
                    $maxConcurrentJobsPerComputer, $jobQueue.Count
                )

                for ($w = 0; $w -lt $workerCount; $w++) {
                    [PSCustomObject]@{
                        Target     = $group.Name
                        JobQueue   = $jobQueue
                        QueueDepth = $jobQueue.Count
                    }
                }
            }

            if ($workers) {
                # busiest computers first so they don't become the critical path
                $workers = @($workers | Sort-Object -Property 'QueueDepth' -Descending)

                Invoke-WithOptionalParallelismHC @params -InputObject $workers
            }
        }

        #region Apply job results to tasks
        foreach ($jobResult in $jobResults) {
            $task = $tasksToExecute[$jobResult.ID]
            $task.Job.Results += $jobResult.Results

            foreach ($jobError in $jobResult.Errors) {
                $task.Job.Errors += $jobError

                Write-Warning (
                    "Job '{0}' failed on '{1}' for '{2}': {3}" -f
                    $task.Type, $task.ComputerName, $task.Path, $jobError
                )
            }
        }
        #endregion
    }
    Catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "$_"
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
}

End {
    try {
        $settings = $file.Settings

        $scriptName = $settings.ScriptName
        $saveInEventLog = $settings.SaveInEventLog
        $sendMail = $settings.SendMail
        $saveLogFiles = $settings.SaveLogFiles

        if (-not $scriptName) {
            Write-Warning "No 'Settings.ScriptName' found in the input file."
            $scriptName = 'Default script name'
        }

        $allLogFilePaths = @()
        $baseLogName = $null
        $logFolderPath = $null

        #region Get job errors
        $jobErrors = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($task in $tasksToExecute) {
            foreach ($jobError in $task.Job.Errors) {
                $M = "{0} on '{1}' with Path '{2}': {3}" -f
                $task.Type, $task.ComputerName, $task.Path, $jobError

                $jobErrors.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = $M
                    }
                )
            }

            foreach ($result in $task.Job.Results.Where({ $_.Error })) {
                $M = "{0} on '{1}' with Path '{2}': {3}" -f
                $task.Type, $task.ComputerName, $result.FullName, $result.Error

                $jobErrors.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = $M
                    }
                )
            }
        }
        #endregion

        #region Counter
        $counter = @{
            removedItems  = @(
                $tasksToExecute.Job.Results | Where-Object { $_.Action -eq 'Removed' }
            ).Count
            removalErrors = @(
                $tasksToExecute.Job.Results | Where-Object { $_.Error }
            ).Count
            jobErrors     = @($tasksToExecute.Job.Errors | Where-Object { $_ }).Count
            systemErrors  = $systemErrors.Count
            totalErrors   = 0
        }
        #endregion

        #region Create log folder
        try {
            $logFolder = Get-StringValueHC $saveLogFiles.Where.Folder

            if ($logFolder) {
                $logFolderPath = Get-LogFolderHC -Path $logFolder

                Write-Verbose "Log folder '$logFolderPath'"

                $baseLogName = Join-Path -Path $logFolderPath -ChildPath (
                    '{0} - {1} ({2})' -f
                    $scriptStartTime.ToString('yyyy_MM_dd_HHmmss'),
                    $scriptName,
                    $(
                        if ($jsonFileItem) { $jsonFileItem.BaseName }
                        else { 'no input file' }
                    )
                )
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed creating the log folder '$($saveLogFiles.Where.Folder)': $_"
                }
            )

            Write-Warning $systemErrors[-1].Message
        }
        #endregion

        $mailParams = @{ }

        $excelParams = @{
            Path               = "$baseLogName - Log.xlsx"
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
            $tasksToExecute
        ) {
            $task.Job.Results | Select-Object -Property 'DateTime',
            'ComputerName',
            'Type',
            @{
                Name       = 'Path'
                Expression = { $_.FullName }
            },
            'CreationTime',
            'LastWriteTime',
            @{
                Name       = 'OlderThan'
                Expression = {
                    if ($task.OlderThan.Unit) {
                        '{0} {1}' -f
                        $task.OlderThan.Quantity, $task.OlderThan.Unit
                    }
                }
            },
            @{
                Name       = 'OlderThanBasedOn'
                Expression = { $task.OlderThan.BasedOn }
            },
            'Action', 'Error'
        }

        if ($excelSheet.Overview -and $baseLogName) {
            Write-Verbose "Export $($excelSheet.Overview.Count) rows to Excel"

            $excelParams.WorksheetName = $excelParams.TableName = 'Overview'

            $excelSheet.Overview | Export-Excel @excelParams

            $allLogFilePaths += $excelParams.Path
        }
        #endregion

        #region Create Excel worksheet Errors
        $excelSheet.Errors += foreach (
            $task in
            $tasksToExecute.where({ $_.Job.Errors })
        ) {
            $task.Job.Errors | Select-Object -Property @{
                Name       = 'ComputerName';
                Expression = { $task.ComputerName }
            },
            @{
                Name       = 'Path';
                Expression = { $task.Path }
            },
            @{
                Name       = 'Type';
                Expression = { $task.Type }
            },
            @{
                Name       = 'OlderThan'
                Expression = {
                    if ($task.OlderThan.Unit) {
                        '{0} {1}' -f
                        $task.OlderThan.Quantity, $task.OlderThan.Unit
                    }
                }
            },
            @{
                Name       = 'Error'
                Expression = { $_ -join ', ' }
            },
            @{
                Name       = 'Stage'
                Expression = { $_.JobStage }
            },
            @{
                Name       = 'TargetObject'
                Expression = { "$($_.TargetObject)" }
            },
            'FullyQualifiedErrorId',
            @{
                Name       = 'ExceptionType'
                Expression = {
                    $exception = $_.Exception
                    if ($exception.SerializedRemoteException) {
                        $exception = $exception.SerializedRemoteException
                    }
                    if ($exception) {
                        $exception.PSTypeNames[0] -replace '^Deserialized\.', ''
                    }
                }
            },
            'ScriptStackTrace',
            @{
                Name       = 'PositionMessage'
                Expression = {
                    if ($_.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
                        $_.Exception.SerializedRemoteInvocationInfo.PositionMessage
                    }
                    else { $_.InvocationInfo.PositionMessage }
                }
            }
        }

        if ($excelSheet.Errors -and $baseLogName) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Errors'

            Write-Verbose (
                "Export {0} rows to sheet '{1}' in Excel file '{2}'" -f
                $excelSheet.Errors.Count,
                $excelParams.WorksheetName, $excelParams.Path
            )

            $excelSheet.Errors | Export-Excel @excelParams

            $allLogFilePaths += $excelParams.Path
        }
        #endregion

        #region Remove old log files
        if (
            ("$($saveLogFiles.DeleteLogsAfterDays)" -match '^\d+$') -and
            ([int]$saveLogFiles.DeleteLogsAfterDays -gt 0) -and
            $logFolderPath
        ) {
            $cutoffDate = (Get-Date).AddDays(-$saveLogFiles.DeleteLogsAfterDays)

            Write-Verbose "Remove log files older than $cutoffDate from '$logFolderPath'"

            Get-ChildItem -LiteralPath $logFolderPath -File |
            Where-Object { $_.LastWriteTime -lt $cutoffDate } |
            ForEach-Object {
                try {
                    $fileToRemove = $_
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                }
                catch {
                    $systemErrors.Add(
                        [PSCustomObject]@{
                            DateTime = Get-Date
                            Message  = "Failed to remove log file '$fileToRemove': $_"
                        }
                    )

                    Write-Warning $systemErrors[-1].Message
                }
            }
        }
        #endregion

        #region Write events to event log
        try {
            if ($saveInEventLog.Save) {
                $eventLogName = Get-StringValueHC $saveInEventLog.LogName

                if (-not $eventLogName) {
                    throw "Both 'Settings.SaveInEventLog.Save' and 'Settings.SaveInEventLog.LogName' are required to save events in the event log."
                }

                @($systemErrors) + @($jobErrors) | ForEach-Object {
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = $_.Message
                            DateTime  = $_.DateTime
                            EntryType = 'Error'
                            EventID   = '2'
                        }
                    )
                }

                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = 'Script ended'
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '199'
                    }
                )

                $params = @{
                    Source  = $scriptName
                    LogName = $eventLogName
                    Events  = $eventLogData
                }
                Write-EventsToEventLogHC @params
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed writing events to the event log: $_"
                }
            )

            Write-Warning $systemErrors[-1].Message
        }
        #endregion

        #region Create system errors log file
        if ($baseLogName -and $systemErrors) {
            $params = @{
            DataToExport   = @($systemErrors)
                PartialPath    = "$baseLogName - System errors log"
                FileExtensions = '.json'
                Append         = $true
            }
            $allLogFilePaths += Out-LogFileHC @params
        }

        $loggedSystemErrorCount = $systemErrors.Count
        #endregion

        $counter.systemErrors = $systemErrors.Count
        $counter.totalErrors = $counter.removalErrors + $counter.jobErrors +
        $counter.systemErrors

        #region Create mail rows
        $mailJobs = foreach ($task in $tasksToExecute) {
            $isUncPath = $task.Path -match '^\\\\([^\\]+)'

            [PSCustomObject]@{
                ComputerName = if ($isUncPath) { $Matches[1] } else { $task.ComputerName }
                Name         = $task.Name
                Path         = $task.Path
                LinkPath     = if ($isUncPath) {
                    $task.Path
                }
                else {
                    $task.Path -replace '^(.):', ('\\{0}\$1$' -f $task.ComputerName)
                }
                Description  = Get-TaskDescriptionHC -Task $task
                Removed      = @(
                    $task.Job.Results | Where-Object { $_.Action -eq 'Removed' }
                ).Count
                Errors       = @(
                    $task.Job.Results | Where-Object { $_.Error }
                ).Count + @($task.Job.Errors | Where-Object { $_ }).Count
            }
        }
        #endregion

        #region Send email
        try {
            $isSendMail = switch ($sendMail.When) {
                'Never' { $false }
                'Always' { $true }
                'OnError' { $counter.totalErrors -gt 0 }
                'OnErrorOrAction' {
                    ($counter.totalErrors -gt 0) -or
                    ($counter.removedItems -gt 0)
                }
                default {
                    throw "Property 'Settings.SendMail.When' with value '$($sendMail.When)' is not supported. Supported values are 'Never', 'Always', 'OnError' or 'OnErrorOrAction'."
                }
            }

            if ($isSendMail) {
                $mailParams += @{
                    From                = Get-StringValueHC $sendMail.From
                    SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
                    SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
                    MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
                    MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
                    Subject             = '{0} removed' -f $counter.removedItems
                    Priority            = 'Normal'
                }

                if ($counter.totalErrors) {
                    $mailParams.Priority = 'High'
                    $mailParams.Subject += ', {0} error{1}' -f
                    $counter.totalErrors,
                    $(if ($counter.totalErrors -ne 1) { 's' })
                }

                if ($sendMail.Subject) {
                    $mailParams.Subject = '{0}, {1}' -f
                    $mailParams.Subject, $sendMail.Subject
                }

                $bodyParams = @{
                    ScriptName      = $scriptName
                    Body            = $sendMail.Body
                    Job             = @($mailJobs | Where-Object { $_ })
                    Removed         = $counter.removedItems
                    Errors          = $counter.totalErrors
                    SystemError     = @($systemErrors.Message)
                    LogFolderPath   = $logFolderPath
                    HasAttachments  = [bool]$allLogFilePaths
                    ScriptStartTime = $scriptStartTime
                }
                $mailParams.Body = Get-MailBodyHtmlHC @bodyParams

                if ($sendMail.FromDisplayName) {
                    $mailParams.FromDisplayName = Get-StringValueHC $sendMail.FromDisplayName
                }

                if ($sendMail.To) {
                    $mailParams.To = $sendMail.To
                }

                if ($sendMail.Bcc) {
                    $mailParams.Bcc = $sendMail.Bcc
                }

                if ($allLogFilePaths) {
                    $mailParams.Attachments = $allLogFilePaths |
                    Sort-Object -Unique
                }

                if ($sendMail.Smtp.ConnectionType) {
                    $mailParams.SmtpConnectionType = Get-StringValueHC $sendMail.Smtp.ConnectionType
                }

                #region Create SMTP credential
                $smtpUserName = Get-StringValueHC $sendMail.Smtp.UserName
                $smtpPassword = Get-StringValueHC $sendMail.Smtp.Password

                if ($smtpUserName -and $smtpPassword) {
                    $securePassword = ConvertTo-SecureString -String $smtpPassword -AsPlainText -Force

                    $mailParams.Credential = New-Object System.Management.Automation.PSCredential(
                        $smtpUserName, $securePassword
                    )
                }
                elseif ($smtpUserName -or $smtpPassword) {
                    throw "Both 'Settings.SendMail.Smtp.UserName' and 'Settings.SendMail.Smtp.Password' are required when authentication is needed."
                }
                #endregion

                Send-MailKitMessageHC @mailParams
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed sending email: $_"
                }
            )

            Write-Warning $systemErrors[-1].Message
        }
        #endregion
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "$_"
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
    finally {
        #region Log system errors that occurred after the log file was created
        $newSystemErrors = @(
            $systemErrors | Select-Object -Skip ([int]$loggedSystemErrorCount)
        )

        if ($newSystemErrors -and $baseLogName) {
            $params = @{
                DataToExport   = $newSystemErrors
                PartialPath    = "$baseLogName - System errors log"
                FileExtensions = '.json'
                Append         = $true
            }
            $null = Out-LogFileHC @params
        }
        #endregion

        Write-Verbose ('Run summary: {0} removed, {1} errors' -f
            ([int]$counter.removedItems), ($systemErrors.Count + $jobErrors.Count))

        if ($systemErrors -or $jobErrors) {
            Write-Warning 'Exit script with error code 1'
            exit 1
        }
        else {
            Write-Verbose 'Script finished successfully'
        }
    }
}