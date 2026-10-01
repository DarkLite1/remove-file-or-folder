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
        path. Relative paths are relative to the script folder.
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
            $null = New-Item -Path $fullPath -ItemType Directory -Force -ErrorAction Stop
        }
        catch {
            throw "Failed creating log folder '$fullPath': $_"
        }
    }

    (Resolve-Path $fullPath).ProviderPath
}

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
