#Requires -Version 5.1
#Requires -Modules ImportExcel

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
    folder  : remove the folder and its contents specified in 'Path'
    content : remove the files in the folder specified in 'Path'

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
            [Boolean]$RemoveEmptyFolders,
            [String]$Name
        )

        Try {
            $compareDate = (Get-Date).AddDays(-$OlderThanDays)
    
            $result = [PSCustomObject]@{
                Type               = $Type
                Name               = $Name
                Path               = $Path
                OlderThanDays      = $OlderThanDays
                OlderThanDate      = $compareDate
                ComputerName       = $env:COMPUTERNAME
                RemoveEmptyFolders = $RemoveEmptyFolders
                Items              = @()
                Error              = $null
            }

            #region Create get params and test file folder
            $commandToRun = "Get-Item -LiteralPath '$Path'"
            $removalType = 'File'

            switch ($Type) {
                'file' {
                    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
                        $result.Items += [PSCustomObject]@{
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
                        $result.Items += [PSCustomObject]@{
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
                    if (-not (Test-Path -LiteralPath $Path -PathType Container)
                    ) {
                        throw 'Folder not found'
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

            [Array]$result.Items = Invoke-Expression $commandToRun | 
            Where-Object { 
                    ($_.CreationTime -lt $compareDate) -or
                    ($OlderThanDays -eq 0)
            } | ForEach-Object {
                try {
                    Remove-Item @removeParams -LiteralPath $_.FullName
                    [PSCustomObject]@{
                        Type         = $removalType
                        FullName     = $_.FullName 
                        CreationTime = $_.CreationTime
                        Action       = 'Removed'
                        Error        = $null
                    }
                }
                catch {
                    [PSCustomObject]@{
                        Type         = $removalType
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
                    $result.Items += $emptyFolders | ForEach-Object {
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
                throw "Input file '$ImportFile' destination path '$($d.Path)': No 'ComputerName' found."
            }
            #endregion

            #region OlderThanDays
            if ($d.PSObject.Properties.Name -notContains 'OlderThanDays') {
                throw "Input file '$ImportFile' destination path '$($d.Path)': Property 'OlderThanDays' not found. Use value number '0' to remove all."
            }
            if (-not ($d.OlderThanDays -is [int])) {
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
        $jobs = @()

        foreach ($d in $Destinations) {
            $invokeParams = @{
                ScriptBlock  = $scriptBlock
                ArgumentList = $d.Remove, $d.Path, $d.OlderThanDays, $d.RemoveEmptyFolders, $d.Name
            }

            $M = "Start job on '{0}' with Remove '{1}' Path '{2}' OlderThanDays '{3}' RemoveEmptyFolders '{4}' Name '{5}'" -f $(
                if ($d.ComputerName) { $d.ComputerName }
                else { $env:COMPUTERNAME }
            ),
            $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
            $invokeParams.ArgumentList[2], $invokeParams.ArgumentList[3], 
            $invokeParams.ArgumentList[4]
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            if ($d.ComputerName) {
                $invokeParams.ComputerName = $d.ComputerName
                $invokeParams.AsJob = $true
                $jobs += Invoke-Command @invokeParams
            }
            else {
                $Jobs += Start-Job @invokeParams
            }
            # & $scriptBlock -Type $d.Remove -Path $d.Path -OlderThanDays $d.OlderThanDays -RemoveEmptyFolders $d.RemoveEmptyFolders
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
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
    Finally {
        Get-Job | Remove-Job
    }
}

End {
    try {
        #region Send mail to user

        #region Error counters
        $removalErrorCount = (
            $jobResults.Items | Where-Object { $_.Error } |
            Measure-Object
        ).Count

        $jobErrorCount = (
            $jobResults | Where-Object { $_.Error } |
            Measure-Object
        ).Count

        $unknownErrorCount = (
            $Error.Exception.Message | Where-Object { $_ } |
            Measure-Object
        ).Count
           
        $totalErrorCount = $removalErrorCount + $jobErrorCount + 
        $unknownErrorCount
        #endregion

        #region Mail subject and priority
        $removedItemsCount = (
            $jobResults.Items | Where-Object { ($_.Action -eq 'Removed') } | Measure-Object
        ).Count

        $mailParams.Subject = "$removedItemsCount removed"

        if ($totalErrorCount) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -lt 1) {
                    's'
                }
            )
        }
        #endregion

        #region Create html lists
        $errorsHtmlList = if ($unknownErrorCount) {
            "<p>During removal <b>$unknownErrorCount non terminating errors</b> were detected:$($Error.Exception.Message | Where-Object { $_ } | ConvertTo-HtmlListHC)</p>"
        }

        $jobResultsHtmlListItems = foreach (
            $job in 
            $jobResults | Sort-Object -Property 'Name', 'Path', 'ComputerName'
        ) {
            "{0}<br>{1}<br>{2}{3}" -f 
            $(
                if ($job.Path -match '^\\\\') {
                    '<a href="{0}">{1}</a>' -f $job.Path, $(
                        if ($job.Name) { $job.Name }
                        else { $job.Path }
                    )
                }
                else {
                    $uncPath = $job.Path -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $job.ComputerName, $job.Path[0]
                    )
                    '<a href="{0}">{1}</a>' -f $uncPath, $(
                        if ($job.Name) { $job.Name }
                        else { $uncPath }
                    )
                }
            ), $(
                $description = if ($job.Type -eq 'File') {
                    if ($job.OlderThanDays -eq 0) {
                        'Remove file'
                    }
                    else {
                        "Remove file when it's older than {0} days" -f 
                        $job.OlderThanDays
                    }
                }
                elseif ($job.Type -eq 'Folder') {
                    if ($job.OlderThanDays -eq 0) {
                        'Remove folder'
                    }
                    else {
                        "Remove folder when it's older than {0} days" -f 
                        $job.OlderThanDays
                    }
                }
                elseif ($job.Type -eq 'Content') {
                    if ($job.OlderThanDays -eq 0) {
                        'Remove folder content'
                    }
                    else {
                        'Remove folder content that is older than {0} days' -f 
                        $job.OlderThanDays
                    }
                }
                if ($job.RemoveEmptyFolders) {
                    $description += ' and remove empty folders'
                }
                $description
            ), $(
                $counters = 'Removed: {0}' -f 
                $(
                    (
                        $job.Items | Where-Object { $_.Action -eq 'Removed' } | 
                        Measure-Object
                    ).Count
                )
                if (
                    $errorCount = (
                        $job.Items | Where-Object { $_.Error } | Measure-Object
                    ).Count
                ) {
                    $counters += ', <b style="color:red;">errors: {0}</b>' -f $errorCount
                }
                $counters
            ), $(
                if ($job.Error) {
                    '<br><b style="color:red;">{0}</b>' -f $job.Error
                }
            )
        }
   
        $jobResultsHtmlList = $jobResultsHtmlListItems | 
        ConvertTo-HtmlListHC -Spacing Wide
        #endregion
        
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "<p>Summary:</p>
                $errorsHtmlList
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