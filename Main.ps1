#Requires -Version 7
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Remove files, folders or folder content on local or remote machines.

.DESCRIPTION
    This script reads a .JSON file containing the paths where files or folders
    need to be removed. When 'OlderThan.Quantity' is '0' all files or folders
    will be removed, depending on the chosen 'Remove' type, regardless their
    creation date.

    All properties of the .JSON file are explained in the '?' section of
    'Example.json'.

.PARAMETER ConfigurationJsonFile
    The path to the .JSON file containing the configuration.

.PARAMETER Path
    The paths to the scripts that execute the removal actions.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ConfigurationJsonFile,
    [HashTable]$Path = @{
        RemoveFileScript          = "$PSScriptRoot\Remove file.ps1"
        RemoveEmptyFoldersScript  = "$PSScriptRoot\Remove empty folders.ps1"
        RemoveFilesInFolderScript = "$PSScriptRoot\Remove files in folder.ps1"
    }
)

Begin {
    $eventLogData = [System.Collections.Generic.List[PSObject]]::new()
    $systemErrors = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date

    function Get-StringValueHC {
        <#
        .SYNOPSIS
            Retrieve a string from the environment variables or a regular
            string.

        .DESCRIPTION
            When the value starts with 'ENV:' the value of that environment
            variable is returned, otherwise the value itself.

        .EXAMPLE
            Get-StringValueHC -Name 'ENV:passwordVariable'

            # Output: the value of $ENV:passwordVariable or an error when the
            # variable does not exist
        #>
        param (
            [String]$Name
        )

        if (-not $Name) {
            return $null
        }
        elseif (
            $Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            $envVariableName = $Name.Substring(4).Trim()
            $envStringValue = Get-Item -Path "Env:\$envVariableName" -EA Ignore
            if ($envStringValue) {
                return $envStringValue.Value
            }
            else {
                throw "Environment variable '$envVariableName' not found."
            }
        }
        else {
            return $Name
        }
    }

    function Invoke-WithOptionalParallelismHC {
        <#
        .SYNOPSIS
            Run a scriptblock for each input object, sequentially or in
            parallel.

        .DESCRIPTION
            With a ThrottleLimit of 1 or less the scriptblock runs in a plain
            foreach loop on the main thread. Otherwise it runs with
            ForEach-Object -Parallel.

            The input object is passed as the first positional argument,
            followed by the values in ArgumentList.

            The scriptblock is rehydrated from its text inside each parallel
            runspace, so '$using:' does not work inside it. Pass everything it
            needs through the input object (DTO) or ArgumentList, and return
            results instead of changing shared objects.
        #>

        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [AllowEmptyCollection()]
            [array]$InputObject,
            [Parameter(Mandatory)]
            [scriptblock]$ScriptBlock,
            [Parameter(Mandatory)]
            [int]$ThrottleLimit,
            [object[]]$ArgumentList = @()
        )

        if ($ThrottleLimit -le 1) {
            foreach ($item in $InputObject) {
                & $ScriptBlock $item @ArgumentList
            }
        }
        else {
            $scriptBlockString = $ScriptBlock.ToString()

            $InputObject | ForEach-Object -Parallel {
                $rehydratedBlock = [scriptblock]::Create($using:scriptBlockString)
                $splatArgs = $using:ArgumentList
                & $rehydratedBlock $_ @splatArgs
            } -ThrottleLimit $ThrottleLimit
        }
    }

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

            @(
                'MaxConcurrent', 'Remove'
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
            throw "Input file '$ConfigurationJsonFile': $_"
        }
        #endregion

        #region Test path exists
        $pathItem = @{}

        $Path.GetEnumerator().ForEach(
            {
                try {
                    $key = $_.Key
                    $value = $_.Value

                    $params = @{
                        Path        = $value
                        ErrorAction = 'Stop'
                    }
                    $pathItem[$key] = (Get-Item @params).FullName
                }
                catch {
                    throw "Path.$key '$value' not found"
                }
            }
        )
        #endregion

        #region Convert .json file
        $PSSessionConfiguration = $file.PSSessionConfiguration

        if (-not $PSSessionConfiguration) {
            $PSSessionConfiguration = 'PowerShell.7'
        }

        $convertScriptBlock = {
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

            #region Add properties
            $_ | Add-Member -NotePropertyMembers @{
                Job = @{
                    Results = @()
                    Errors  = @()
                }
            }
            #endregion
        }
        #endregion

        #region Create tasks to execute
        $tasksToExecute = @()

        $file.Remove.File.foreach(
            {
                & $convertScriptBlock

                $tasksToExecute += $_ | Select-Object -Property *,
                @{
                    Name       = 'Type'
                    Expression = { 'RemoveFile' }
                }
            }
        )

        $file.Remove.FilesInFolder.foreach(
            {
                & $convertScriptBlock

                $tasksToExecute += $_ | Select-Object -Property *,
                @{
                    Name       = 'Type'
                    Expression = { 'RemoveFilesInFolder' }
                }
            }
        )

        $file.Remove.EmptyFolders.foreach(
            {
                & $convertScriptBlock

                $tasksToExecute += $_ | Select-Object -Property *,
                @{
                    Name       = 'Type'
                    Expression = { 'RemoveEmptyFolders' }
                }
            }
        )

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
                    $filePath = $pathItem.RemoveFileScript
                    $argumentList = @(
                        $task.Path, $task.OlderThan.Unit, $task.OlderThan.Quantity
                    )

                    $M = "Start job '$_' on '{0}' with Path '{1}' OlderThan.Quantity '{3}' OlderThan.Unit '{2}'" -f
                    $task.ComputerName,
                    $argumentList[0], $argumentList[1], $argumentList[2]

                    break
                }
                'RemoveFilesInFolder' {
                    $filePath = $pathItem.RemoveFilesInFolderScript
                    $argumentList = @(
                        $task.Path, $task.OlderThan.Unit, $task.OlderThan.Quantity, $task.Recurse
                    )

                    $M = "Start job '$_' on '{0}' with Path '{1}' OlderThan.Quantity '{3}' OlderThan.Unit '{2}' Recurse '{4}'" -f
                    $task.ComputerName,
                    $argumentList[0], $argumentList[1], $argumentList[2],
                    $argumentList[3]

                    break
                }
                'RemoveEmptyFolders' {
                    $filePath = $pathItem.RemoveEmptyFoldersScript
                    $argumentList = @($task.Path)

                    $M = "Start job '$_' on '{0}' with Path '{1}'" -f
                    $task.ComputerName, $argumentList[0]

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
                FilePath     = $filePath
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
                [Int]$RetryDelaySeconds
            )

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
                        if ($dto.ComputerName -eq $env:COMPUTERNAME) {
                            $arguments = $dto.ArgumentList
                            $result.Results = @(& $dto.FilePath @arguments)
                        }
                        else {
                            $sessionParams = @{
                                ComputerName      = $dto.ComputerName
                                ConfigurationName = $SessionConfiguration
                                SessionOption     = $sessionOption
                                ErrorAction       = 'Stop'
                            }
                            $session = New-PSSession @sessionParams

                            $invokeParams = @{
                                Session      = $session
                                FilePath     = $dto.FilePath
                                ArgumentList = $dto.ArgumentList
                                ErrorAction  = 'Stop'
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
            ArgumentList  = $PSSessionConfiguration, 3, 5
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
                    'Error for {0} : {1}' -f
                    @($taskDtos)[$jobResult.ID].StartMessage, $jobError
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
    function ConvertTo-HtmlListHC {
        <#
        .SYNOPSIS
            Creates an unordered HTML list.

        .EXAMPLE
            'Item 1', 'Item 2' | ConvertTo-HtmlListHC

            Creates '<ul><li style="margin: 10px 0;">Item 1</li>..</ul>'
        #>
        param (
            [parameter(Mandatory, ValueFromPipeline)]
            [String[]]$Message,
            [String]$Header,
            [String]$FootNote
        )

        begin {
            $allItems = [System.Collections.ArrayList]::new()
        }

        process {
            $null = $allItems.AddRange($Message)
        }

        end {
            @"
$($Header ? "<h3>$Header</h3>" : '')
<ul>
    $(
        $allItems |
        ForEach-Object { "<li style=`"margin: 10px 0;`">$_</li>" }
    )
</ul>
$($FootNote ? "<i><font size=`"2`">* $FootNote</font></i>" : '')
"@
        }
    }

    function Get-LogFolderHC {
        <#
        .SYNOPSIS
            Create the log folder when it doesn't exist and return its full
            path. Relative paths are relative to $PSScriptRoot.
        #>

        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if ($Path -match '^[a-zA-Z]:\\' -or $Path -match '^\\') {
            $fullPath = $Path
        }
        else {
            $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $Path
        }

        if (-not (Test-Path -Path $fullPath -PathType Container)) {
            try {
                Write-Verbose "Create log folder '$fullPath'"
                $null = New-Item -Path $fullPath -ItemType Directory -Force
            }
            catch {
                throw "Failed creating log folder '$fullPath': $_"
            }
        }

        (Resolve-Path $fullPath).ProviderPath
    }

    function Out-LogFileHC {
        <#
        .SYNOPSIS
            Export objects to a .json or .txt log file.
        #>

        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [PSCustomObject[]]$DataToExport,
            [Parameter(Mandatory)]
            [String]$PartialPath,
            [Parameter(Mandatory)]
            [ValidateSet('.json', '.txt')]
            [String[]]$FileExtensions,
            [Switch]$Append
        )

        $allLogFilePaths = @()

        foreach (
            $fileExtension in
            $FileExtensions | Sort-Object -Unique
        ) {
            try {
                $logFilePath = "$PartialPath{0}" -f $fileExtension

                Write-Verbose (
                    "Export {0} object{1} to '$logFilePath'" -f
                    $DataToExport.Count,
                    $(if ($DataToExport.Count -ne 1) { 's' })
                )

                switch ($fileExtension) {
                    '.json' {
                        $convertedDataToExport = foreach (
                            $exportObject in
                            $DataToExport
                        ) {
                            [PSCustomObject]@{
                                DateTime = $exportObject.DateTime
                                Message  = "$($exportObject.Message)"
                            }
                        }

                        if (
                            $Append -and
                            (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                        ) {
                            $params = @{
                                LiteralPath = $logFilePath
                                Raw         = $true
                                Encoding    = 'UTF8'
                            }
                            $jsonFileContent = Get-Content @params | ConvertFrom-Json

                            $convertedDataToExport = [array]$convertedDataToExport + [array]$jsonFileContent
                        }

                        $convertedDataToExport |
                        ConvertTo-Json -Depth 7 |
                        Out-File -LiteralPath $logFilePath

                        break
                    }
                    '.txt' {
                        $DataToExport | Format-List -Property * -Force |
                        Out-File -LiteralPath $logFilePath -Append:$Append

                        break
                    }
                }

                $allLogFilePaths += $logFilePath
            }
            catch {
                Write-Warning "Failed creating log file '$logFilePath': $_"
            }
        }

        $allLogFilePaths
    }

    function Send-MailKitMessageHC {
        <#
        .SYNOPSIS
            Send an email using MailKit and MimeKit assemblies.

        .DESCRIPTION
            Requires the assemblies to be installed:

            $params = @{
                Source           = 'https://www.nuget.org/api/v2'
                SkipDependencies = $true
                Scope            = 'AllUsers'
            }
            Install-Package @params -Name 'MailKit'
            Install-Package @params -Name 'MimeKit'
        #>

        [CmdletBinding()]
        param (
            [parameter(Mandatory)]
            [string]$MailKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$MimeKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$SmtpServerName,
            [parameter(Mandatory)]
            [ValidateSet(25, 465, 587, 2525)]
            [int]$SmtpPort,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
            [parameter(Mandatory)]
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [string]$FromDisplayName,
            [string[]]$To,
            [string[]]$Bcc,
            [int]$MaxAttachmentSize = 20MB,
            [ValidateSet(
                'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
            )]
            [string]$SmtpConnectionType = 'None',
            [ValidateSet('Normal', 'Low', 'High')]
            [string]$Priority = 'Normal',
            [string[]]$Attachments,
            [PSCredential]$Credential
        )

        begin {
            function Test-IsAssemblyLoaded {
                param (
                    [String]$Name
                )
                foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                    if ($assembly.FullName -like "$Name, Version=*") {
                        return $true
                    }
                }
                return $false
            }

            function Add-Attachments {
                param (
                    [string[]]$Attachments,
                    [MimeKit.Multipart]$BodyMultiPart
                )

                $attachmentList = New-Object System.Collections.ArrayList($null)

                foreach (
                    $attachmentPath in
                    $Attachments | Sort-Object -Unique
                ) {
                    try {
                        try {
                            $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                            if ($attachmentItem.PSIsContainer) {
                                Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                                continue
                            }
                        }
                        catch {
                            Write-Warning "Attachment '$attachmentPath' not found"
                            continue
                        }

                        $totalSizeAttachments += $attachmentItem.Length

                        $null = $attachmentList.Add($attachmentItem)

                        if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                            $M = 'The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details.' -f
                            ([math]::Round(($MaxAttachmentSize / 1MB))),
                            ([math]::Round(($totalSizeAttachments / 1MB), 2))

                            Write-Warning $M

                            return [PSCustomObject]@{
                                AttachmentLimitExceededMessage = $M
                            }
                        }
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentPath': $_"
                    }
                }

                foreach (
                    $attachmentItem in
                    $attachmentList
                ) {
                    try {
                        Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                        $attachment = New-Object MimeKit.MimePart

                        $memoryStream = New-Object System.IO.MemoryStream

                        try {
                            $fileStream = [System.IO.File]::OpenRead($attachmentItem.FullName)
                            $fileStream.CopyTo($memoryStream)
                        }
                        finally {
                            if ($fileStream) {
                                $fileStream.Dispose()
                            }
                        }

                        $memoryStream.Position = 0

                        $attachment.Content = New-Object MimeKit.MimeContent($memoryStream)

                        $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                        $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                        $attachment.FileName = $attachmentItem.Name

                        $bodyMultiPart.Add($attachment)
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentItem': $_"
                    }
                }
            }

            try {
                if (-not ($To -or $Bcc)) {
                    throw "Either 'To' to 'Bcc' is required for sending emails"
                }

                foreach ($email in $To) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "To email address '$email' not valid."
                    }
                }

                foreach ($email in $Bcc) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "Bcc email address '$email' not valid."
                    }
                }

                if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                    try {
                        Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                        Add-Type -Path $MimeKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                    }
                }

                if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                    try {
                        Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                        Add-Type -Path $MailKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                    }
                }
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
        }

        process {
            try {
                $message = New-Object -TypeName 'MimeKit.MimeMessage'

                $bodyPart = New-Object MimeKit.TextPart('html')
                $bodyPart.Text = $Body

                $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
                $bodyMultiPart.Add($bodyPart)

                if ($Attachments) {
                    $params = @{
                        Attachments   = $Attachments
                        BodyMultiPart = $bodyMultiPart
                    }
                    $addAttachments = Add-Attachments @params

                    if ($addAttachments.AttachmentLimitExceededMessage) {
                        $bodyPart.Text += '<p><i>{0}</i></p>' -f
                        $addAttachments.AttachmentLimitExceededMessage
                    }
                }

                $message.Body = $bodyMultiPart

                $fromAddress = New-Object MimeKit.MailboxAddress(
                    $FromDisplayName, $From
                )
                $message.From.Add($fromAddress)

                foreach ($email in $To) {
                    $message.To.Add($email)
                }

                foreach ($email in $Bcc) {
                    $message.Bcc.Add($email)
                }

                $message.Subject = $Subject

                switch ($Priority) {
                    'Low' {
                        $message.Headers.Add('X-Priority', '5 (Lowest)')
                        break
                    }
                    'Normal' {
                        $message.Headers.Add('X-Priority', '3 (Normal)')
                        break
                    }
                    'High' {
                        $message.Headers.Add('X-Priority', '1 (Highest)')
                        break
                    }
                    default {
                        throw "Priority type '$_' not supported"
                    }
                }

                $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

                try {
                    $smtp.Connect(
                        $SmtpServerName, $SmtpPort,
                        [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
                    )
                }
                catch {
                    throw "Failed to connect to SMTP server '$SmtpServerName' on port '$SmtpPort' with connection type '$SmtpConnectionType': $_"
                }

                if ($Credential) {
                    try {
                        $smtp.Authenticate(
                            $Credential.UserName,
                            $Credential.GetNetworkCredential().Password
                        )
                    }
                    catch {
                        throw "Failed to authenticate with user name '$($Credential.UserName)' to SMTP server '$SmtpServerName': $_"
                    }
                }

                Write-Verbose "Send mail to '$To' with subject '$Subject'"

                $null = $smtp.Send($message)
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
            finally {
                if ($smtp) {
                    $smtp.Disconnect($true)
                    $smtp.Dispose()
                }
                if ($message) {
                    $message.Dispose()
                }
            }
        }
    }

    function Write-EventsToEventLogHC {
        <#
        .SYNOPSIS
            Write events to the event log.

        .DESCRIPTION
            Custom EventID's based on the PowerShell streams:
            100 Script started, 4 Verbose, 1 Output, 3 Warning, 2 Error,
            199 Script ended.

            All properties of an event that are not 'EntryType' or 'EventID'
            are used to create the message.
        #>

        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [String]$Source,
            [Parameter(Mandatory)]
            [String]$LogName,
            [PSCustomObject[]]$Events
        )

        try {
            if ([System.Diagnostics.EventLog]::SourceExists($Source)) {
                $existingLogName = [System.Diagnostics.EventLog]::LogNameFromSourceName($Source, '.')

                if ($existingLogName -ne $LogName) {
                    throw "The event log source '$Source' is already registered with event log name '$existingLogName', it cannot be used with log name '$LogName'."
                }
            }
            else {
                Write-Verbose "Create event log source '$Source' with log name '$LogName'"

                New-EventLog -LogName $LogName -Source $Source -EA Stop
            }

            foreach ($eventItem in $Events) {
                $params = @{
                    LogName     = $LogName
                    Source      = $Source
                    EntryType   = $eventItem.EntryType
                    EventID     = $eventItem.EventID
                    Message     = ''
                    ErrorAction = 'Stop'
                }

                if (-not $params.EntryType) {
                    $params.EntryType = 'Information'
                }
                if (-not $params.EventID) {
                    $params.EventID = 4
                }

                foreach (
                    $property in
                    $eventItem.PSObject.Properties | Where-Object {
                        ($_.Name -ne 'EntryType') -and ($_.Name -ne 'EventID')
                    }
                ) {
                    $params.Message += "`n- $($property.Name) '$($property.Value)'"
                }

                Write-EventLog @params
            }
        }
        catch {
            throw "Failed to write to event log '$LogName' source '$Source': $_"
        }
    }

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
            @{
                Name       = 'OlderThan'
                Expression = {
                    if ($task.OlderThan.Unit) {
                        '{0} {1}' -f
                        $task.OlderThan.Quantity, $task.OlderThan.Unit
                    }
                }
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
        $errorsToLog = @($systemErrors) + @($jobErrors)

        if ($baseLogName -and $errorsToLog) {
            $params = @{
                DataToExport   = $errorsToLog
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

        #region Create html lists
        $systemErrorsHtmlList = if ($counter.systemErrors) {
            '<p>Detected <b>{0} system error{1}</b>:{2}</p>' -f
            $counter.systemErrors,
            $(if ($counter.systemErrors -ne 1) { 's' }),
            $($systemErrors.Message | Where-Object { $_ } | ConvertTo-HtmlListHC)
        }

        $jobResultsHtmlListItems = foreach (
            $task in
            $tasksToExecute |
            Sort-Object -Property 'Name', 'Path', 'ComputerName'
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
                $description = switch ($task.Type) {
                    'RemoveFile' {
                        'Remove file'
                        break
                    }
                    'RemoveFilesInFolder' {
                        'Remove files in folder'
                        break
                    }
                    'RemoveEmptyFolders' {
                        'Remove empty folders'
                        break
                    }
                    Default {
                        throw "Type '$_' not supported"
                    }
                }

                if ($task.OlderThan.Quantity) {
                    $description += ' older than {0} {1}{2}' -f
                    $($task.OlderThan.Quantity),
                    $($task.OlderThan.Unit.ToLower()),
                    $(
                        if ($task.OlderThan.Quantity -ne 1) { 's' }
                    )
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

        $jobResultsHtmlList = if ($jobResultsHtmlListItems) {
            $jobResultsHtmlListItems | ConvertTo-HtmlListHC
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

                $mailParams.Body = @"
<!DOCTYPE html>
<html>
<head>
<style type="text/css">
    body {
        font-family:verdana;
        font-size:14px;
        background-color:white;
    }
    h1 {
        margin-bottom: 0;
    }
    table {
        border-collapse:collapse;
        border:0px none;
        padding:3px;
        text-align:left;
    }
    td, th {
        border-collapse:collapse;
        border:1px none;
        padding:3px;
        text-align:left;
    }
    #aboutTable th {
        color: rgb(143, 140, 140);
        font-weight: normal;
    }
    #aboutTable td {
        color: rgb(143, 140, 140);
        font-weight: normal;
    }
</style>
</head>
<body>
<table>
    <h1>$scriptName</h1>
    <hr size="2" color="#06cc7a">

    $($sendMail.Body)

    <table>
        <tr>
            <th>Removed</th>
            <td>$($counter.removedItems)</td>
        </tr>
        $(
            $counter.removalErrors ?
            "<tr style=`"background-color: #ffe5ec;`">
                <th>Removal errors</th>
                <td>$($counter.removalErrors)</td>
            </tr>" : ''
        )
        $(
            $counter.jobErrors ?
            "<tr style=`"background-color: #ffe5ec;`">
                <th>Job errors</th>
                <td>$($counter.jobErrors)</td>
            </tr>" : ''
        )
        $(
            $counter.systemErrors ?
            "<tr style=`"background-color: #ffe5ec;`">
                <th>System errors</th>
                <td>$($counter.systemErrors)</td>
            </tr>" : ''
        )
    </table>

    $systemErrorsHtmlList

    $(if ($jobResultsHtmlList) { "<p>Summary:</p>$jobResultsHtmlList" })

    $(
        if ($allLogFilePaths) {
            '<p><i>* Check the attachment(s) for details</i></p>'
        }
    )

    <hr size="2" color="#06cc7a">
    <table id="aboutTable">
        $(
            '<tr>
                <th>Start time</th>
                <td>{0:00}/{1:00}/{2:00} {3:00}:{4:00} ({5})</td>
            </tr>' -f
            $scriptStartTime.Day,
            $scriptStartTime.Month,
            $scriptStartTime.Year,
            $scriptStartTime.Hour,
            $scriptStartTime.Minute,
            $scriptStartTime.DayOfWeek
        )
        $(
            $runTime = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
            '<tr>
                <th>Duration</th>
                <td>{0:00}:{1:00}:{2:00}</td>
            </tr>' -f
            $runTime.Hours, $runTime.Minutes, $runTime.Seconds
        )
        $(
            if ($logFolderPath) {
                '<tr>
                    <th>Log files</th>
                    <td><a href="{0}">Open log folder</a></td>
                </tr>' -f $logFolderPath
            }
        )
        <tr>
            <th>Host</th>
            <td>$($host.Name)</td>
        </tr>
        <tr>
            <th>PowerShell</th>
            <td>$($PSVersionTable.PSVersion.ToString())</td>
        </tr>
        <tr>
            <th>Computer</th>
            <td>$env:COMPUTERNAME</td>
        </tr>
        <tr>
            <th>Account</th>
            <td>$env:USERDNSDOMAIN\$env:USERNAME</td>
        </tr>
    </table>
</table>
</body>
</html>
"@

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

        if ($systemErrors -or $jobErrors) {
            Write-Warning 'Exit script with error code 1'
            exit 1
        }
        else {
            Write-Verbose 'Script finished successfully'
        }
    }
}