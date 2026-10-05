<#
.SYNOPSIS
    Shared email, logging and execution helpers for Main.ps1.

.DESCRIPTION
    Dot-source this file to load its functions. Loading it does not run
    removal tasks, send email or create log files.

.EXAMPLE
    . '.\Functions.ps1'

    Makes the helpers available in the current scope.
#>


function Get-MailThemeHC {
    <#
    .SYNOPSIS
        Colors and fonts of the e-mail, the same palette as the Permission
        matrix script.
    #>
    @{
        StatusError   = '#fee2e2'
        AccentError   = '#dc2626'
        AccentSuccess = '#16a34a'
        AccentIdle    = '#6b7280'
        AccentSystem  = '#7c2d12'
        GradError     = @('#7f1d1d', '#dc2626')
        GradSuccess   = @('#14532d', '#16a34a')
        GradIdle      = @('#374151', '#6b7280')
        TextMain      = '#111827'
        TextMuted     = '#374151'
        TextLight     = '#6b7280'
        BgPage        = '#e5e7eb'
        BgWhite       = '#ffffff'
        BorderMain    = '#d1d5db'
        BorderLight   = '#e5e7eb'
        LinkColor     = '#2563eb'
        FontStack     = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif"
        MonoStack     = 'Consolas, Menlo, monospace'
        BodyWidth     = 620
    }
}

function ConvertTo-FileUrlHC {
    <#
    .SYNOPSIS
        Convert a Windows path (UNC or local) to a 'file://' URL for an href.
    #>
    param ([String]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    'file://' + ($Path -replace '\\', '/' -replace ' ', '%20')
}

function New-PillHtmlHC {
    <#
    .SYNOPSIS
        A colored status pill.

    .DESCRIPTION
        Browsers get a rounded span. Outlook (Word engine) ignores
        border-radius, so it gets a VML roundrect instead. Conditional
        comments make sure every client renders exactly one of them.
    #>
    param (
        [String]$Text,
        [String]$Bg,
        [String]$Color = '#ffffff'
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $span = "<span style=`"display:inline-block; padding:3px 10px; background-color:$Bg; color:$Color; border-radius:12px; font-size:11px; font-weight:700; letter-spacing:0.3px; text-transform:uppercase; line-height:1.6;`">$Text</span>"

    # VML needs a fixed width, estimated from the upper-cased text
    $upper = $Text.ToUpper()
    $width = [int][Math]::Ceiling(($upper.Length * 8.5) + 26)

    $vml = "<v:roundrect xmlns:v=`"urn:schemas-microsoft-com:vml`" xmlns:w=`"urn:schemas-microsoft-com:office:word`" arcsize=`"50%`" fillcolor=`"$Bg`" stroked=`"f`" style=`"height:26px; width:${width}px; v-text-anchor:middle; mso-padding-alt:0;`">" +
    '<w:anchorlock/>' +
    "<center style=`"color:$Color; font-family:sans-serif; font-size:11px; font-weight:700; letter-spacing:0.3px;`">$upper</center>" +
    '</v:roundrect>'

    "<!--[if mso]>$vml<![endif]--><!--[if !mso]><!-->$span<!--<![endif]-->"
}

function Get-TaskDescriptionHC {
    <#
    .SYNOPSIS
        A readable description of what a task does, for the e-mail.

    .EXAMPLE
        Get-TaskDescriptionHC -Task ([pscustomobject]@{
            Type = 'RemoveFilesInFolder'
            OlderThan = @{ Quantity = 3; Unit = 'Month'; BasedOn = 'LastWriteTime' }
            Recurse = $true
        })

        Returns 'Remove files older than 3 months (last write time), including subfolders'.

    .PARAMETER Task
        An internal task created by Main.ps1, not a raw JSON task. Type is
        RemoveFile, RemoveFilesInFolder or RemoveEmptyFolders. Optional
        ExcludeFolders and ExcludeFiles lists add exclusion counts.
    #>
    param (
        [Parameter(Mandatory)]
        [PSCustomObject]$Task
    )

    $olderThan = ''

    if ($Task.OlderThan -and ([int]$Task.OlderThan.Quantity -gt 0)) {
        $quantity = [int]$Task.OlderThan.Quantity
        $basedOn = switch ($Task.OlderThan.BasedOn) {
            'LastWriteTime' { 'last write time' }
            'CreationTime' { 'creation time' }
            default { $_ }
        }

        $olderThan = ' older than {0} {1}{2} ({3})' -f
        $quantity, $Task.OlderThan.Unit.ToLower(),
        $(if ($quantity -ne 1) { 's' }), $basedOn
    }

    $description = switch ($Task.Type) {
        'RemoveFile' {
            "Remove file$olderThan"
        }
        'RemoveFilesInFolder' {
            $text = if ($olderThan) { "Remove files$olderThan" } else { 'Remove all files' }
            if ($Task.Recurse) { $text += ', including subfolders' }
            $text
        }
        'RemoveEmptyFolders' {
            'Remove empty folders'
        }
        default {
            throw "Type '$_' not supported"
        }
    }

    $excluded = @()
    $folderCount = @($Task.ExcludeFolders | Where-Object { $_ }).Count
    $fileCount = @($Task.ExcludeFiles | Where-Object { $_ }).Count

    if ($folderCount) {
        $excluded += '{0} folder{1}' -f $folderCount, $(if ($folderCount -ne 1) { 's' })
    }
    if ($fileCount -and ($Task.Type -eq 'RemoveFilesInFolder')) {
        $excluded += '{0} file{1}' -f $fileCount, $(if ($fileCount -ne 1) { 's' })
    }
    if ($excluded) {
        $description += ', excluding ' + ($excluded -join ' and ')
    }

    $description
}

function Build-MailSummaryBannerHC {
    <#
    .SYNOPSIS
        The line with the totals as colored pills, at the top of the e-mail.
    #>
    param (
        [Int]$Removed,
        [Int]$Errors
    )

    $theme = Get-MailThemeHC

    # Outlook hides the VML label in a cell without these settings
    $pillTd = {
        param ([String]$Text, [String]$Bg)
        "<td valign='middle' style='vertical-align:middle; padding:4px 6px 4px 0; white-space:nowrap; font-size:0;'>$(New-PillHtmlHC -Text $Text -Bg $Bg)</td>"
    }

    $pills = @(
        & $pillTd "$Removed Removed" $(
            if ($Removed) { $theme.AccentSuccess } else { $theme.AccentIdle }
        )
    )
    if ($Errors) {
        $pills += & $pillTd (
            "$Errors Error" + $(if ($Errors -ne 1) { 's' })
        ) $theme.AccentError
    }

    @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 16px 0;">
    <tr>
        <td style='padding:0;'>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
                <tr>
                    <td valign='middle' style='vertical-align:middle; padding:4px 12px 4px 0; font-size:13px; font-weight:600; color:$($theme.TextMain);'>Summary</td>
                    $($pills -join '')
                </tr>
            </table>
        </td>
    </tr>
</table>
<!--[if mso]>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td height="16" style="font-size:0; line-height:16px; mso-line-height-rule:exactly;">&#160;</td></tr></table>
<![endif]-->
"@
}

function Build-MailSystemErrorsBlockHC {
    <#
    .SYNOPSIS
        A red card per system error, shown above the task results.
    #>
    param (
        [String[]]$Message
    )

    $Message = @($Message | Where-Object { $_ })
    if (-not $Message) { return '' }

    $theme = Get-MailThemeHC
    $pill = New-PillHtmlHC -Text 'System Error' -Bg $theme.AccentSystem

    $rows = foreach ($text in $Message) {
        $encoded = [System.Net.WebUtility]::HtmlEncode($text)

        @"
<tr>
    <td style='padding:0 0 8px 0;'>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="$($theme.StatusError)" style="border-collapse:separate; width:100%; background-color:$($theme.StatusError); border-left:3px solid $($theme.AccentError); border-radius:6px;">
            <tr>
                <td valign="middle" width="26" style='padding:12px 0 12px 14px; color:$($theme.AccentError); font-size:16px; font-weight:bold; line-height:1;'>&#10006;</td>
                <td valign="middle" style='padding:10px 12px; color:$($theme.TextMuted); font-size:12px; line-height:1.5; font-family:$($theme.MonoStack); overflow-wrap:anywhere; word-break:break-word;'>$encoded</td>
                <td valign="middle" align="right" style='padding:10px 14px 10px 6px; white-space:nowrap; font-size:0;'>$pill</td>
            </tr>
        </table>
    </td>
</tr>
"@
    }

    $label = 'System Errors ({0})' -f $Message.Count

    @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 20px 0; table-layout:fixed; width:100%;">
    <tr>
        <td style='padding:0 0 8px 0; font-size:11px; font-weight:700; color:$($theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'><p style='margin:0; mso-line-height-rule:exactly; line-height:14px;'>$label</p></td>
    </tr>
    $($rows -join '')
</table>
<!--[if mso]>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td height="12" style="font-size:0; line-height:12px; mso-line-height-rule:exactly;">&#160;</td></tr></table>
<![endif]-->
"@
}

function Build-MailJobRowHC {
    <#
    .SYNOPSIS
        One row per input task in a computer card: what was done and the result.

    .PARAMETER Job
        Object containing Entries, Description, Removed and Errors. Each entry
        contains Name, Path and LinkPath. A single-path job can supply those
        properties directly. Name is optional and appears above Path.
        Removed and Errors are counts. Text is HTML-encoded before rendering.
    #>
    param (
        [Parameter(Mandatory)]
        [PSCustomObject]$Job
    )

    $theme = Get-MailThemeHC

    $accent, $pill = if ($Job.Errors) {
        $theme.AccentError, (New-PillHtmlHC -Text 'Error' -Bg $theme.AccentError)
    }
    elseif ($Job.Removed) {
        $theme.AccentSuccess, ''
    }
    else {
        $theme.AccentIdle, ''
    }

    $description = [System.Net.WebUtility]::HtmlEncode($Job.Description)

    $entries = if ($Job.Entries) { $Job.Entries } else { @($Job) }
    $titleHtml = (@(foreach ($entry in $entries) {
        $href = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $entry.LinkPath))
        $path = [System.Net.WebUtility]::HtmlEncode($entry.Path)

        if ($entry.Name) {
            "<div style='margin:0; font-weight:700; color:$($theme.TextMain); font-size:13px; line-height:16px; mso-line-height-rule:exactly;'><a href='$href' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$($theme.TextMain);'>$([System.Net.WebUtility]::HtmlEncode($entry.Name))</a></div>" +
            "<div style='margin:0; font-family:$($theme.MonoStack); font-size:11px; color:$($theme.TextMuted); line-height:14px; mso-line-height-rule:exactly; overflow-wrap:anywhere; word-break:break-all;'>$path</div>"
        }
        else {
            "<div style='margin:0; font-family:$($theme.MonoStack); font-weight:700; font-size:12px; color:$($theme.TextMain); line-height:16px; mso-line-height-rule:exactly; overflow-wrap:anywhere; word-break:break-all;'><a href='$href' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$($theme.TextMain);'>$path</a></div>"
        }
    })) -join ''

    $resultText = '{0} removed' -f $Job.Removed
    if ($Job.Errors) {
        $resultText += '<br>{0} error{1}' -f $Job.Errors, $(if ($Job.Errors -ne 1) { 's' })
    }

    @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; width:100%; margin:0 0 4px 0; table-layout:fixed; background-color:$($theme.BgWhite); border:1px solid $($theme.BorderMain); border-left:3px solid $accent; border-radius:6px;">
    <tr>
        <td valign='middle' width='20' style='vertical-align:middle; padding:6px 0 6px 12px; color:$accent; font-size:12px; line-height:15px; mso-line-height-rule:exactly;'>&#9679;</td>
        <td valign='middle' style='vertical-align:middle; padding:6px 8px;'>
            $titleHtml
            <div style='margin:2px 0 0 0; font-size:11px; color:$($theme.TextLight); line-height:14px; mso-line-height-rule:exactly;'>$description</div>
        </td>
        <td valign='middle' align='right' nowrap='nowrap' width='76' style='vertical-align:middle; padding:6px 8px; color:$($theme.TextMuted); font-size:11px; line-height:15px; mso-line-height-rule:exactly; white-space:nowrap; text-align:right;'>$resultText</td>
        <td valign='middle' align='right' width='70' style='vertical-align:middle; padding:4px 12px 4px 4px; white-space:nowrap; font-size:0;'>$(if ($pill) { $pill } else { '&nbsp;' })</td>
    </tr>
</table>
"@
}

function Build-MailComputerCardHC {
    <#
    .SYNOPSIS
        A card per computer or computer set: a header and a row per input task.

    .DESCRIPTION
        The header is red when a job failed, green when something was
        removed and grey when nothing was removed. Outlook cannot render the
        gradient, so it gets the average color of the two gradient stops.
    #>
    param (
        [Parameter(Mandatory)]
        [String]$ComputerName,
        [Parameter(Mandatory)]
        [PSCustomObject[]]$Job
    )

    $theme = Get-MailThemeHC

    $removed = ($Job | Measure-Object -Property Removed -Sum).Sum
    $errors = ($Job | Measure-Object -Property Errors -Sum).Sum

    $symbol, $gradient = if ($errors) {
        '&#10006;', $theme.GradError
    }
    elseif ($removed) {
        '&#10003;', $theme.GradSuccess
    }
    else {
        '&#8211;', $theme.GradIdle
    }
    $gradFrom, $gradTo = $gradient

    $gradMid = '#' + ((0, 2, 4 | ForEach-Object {
                '{0:x2}' -f [int][Math]::Round((
                        [Convert]::ToInt32($gradFrom.Substring(1 + $_, 2), 16) +
                        [Convert]::ToInt32($gradTo.Substring(1 + $_, 2), 16)
                    ) / 2)
            }) -join '')

    $headerLabel = '{0}&nbsp;removed' -f $removed
    if ($errors) {
        $headerLabel += ' &middot; {0}&nbsp;error{1}' -f $errors, $(if ($errors -ne 1) { 's' })
    }

    # rows with errors first, then rows that removed something
    $rows = ($Job | Sort-Object -Property @{
            Expression = { if ($_.Errors) { 0 } elseif ($_.Removed) { 1 } else { 2 } }
        }, Path, Description | ForEach-Object {
            Build-MailJobRowHC -Job $_
        }) -join '<!--[if mso]><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff"><tr><td bgcolor="#ffffff" height="4" style="font-size:0; line-height:4px; mso-line-height-rule:exactly;">&#160;</td></tr></table><![endif]-->'

    $pathCount = @($Job | ForEach-Object {
        if ($_.Entries) { $_.Entries.Path } else { $_.Path }
    } | Sort-Object -Unique).Count

    @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="$($theme.BgWhite)" style="border-collapse:separate; margin:0 0 16px 0; table-layout:fixed; width:100%; background-color:$($theme.BgWhite); border:1px solid $($theme.BorderLight); border-radius:10px; overflow:hidden; box-shadow:0 2px 4px rgba(0,0,0,0.06);">
    <tr>
        <td bgcolor="$gradMid" style='padding:0; background-color:$gradMid; background-image:linear-gradient(135deg, $gradFrom 0%, $gradTo 100%);'>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                <tr>
                    <td valign='middle' align='center' width='52' style='vertical-align:middle; text-align:center; padding:12px 0; font-size:18px; font-weight:bold; color:#ffffff; line-height:24px; mso-line-height-rule:exactly;'>$symbol</td>
                    <td valign='middle' style='padding:12px 8px 12px 0;'>
                        <p style='margin:0; font-size:16px; font-weight:700; color:#ffffff; line-height:20px; mso-line-height-rule:exactly;'>$([System.Net.WebUtility]::HtmlEncode($ComputerName))</p>
                        <p style='margin:2px 0 0 0; font-size:12px; color:#f1f2f4; line-height:17px; mso-line-height-rule:exactly;'>$pathCount path$(if ($pathCount -ne 1) { 's' })</p>
                    </td>
                    <td valign='middle' align='right' width='140' style='padding:12px 14px 12px 6px; white-space:nowrap; font-size:12px; font-weight:700; color:#e5e7eb; text-transform:uppercase; letter-spacing:0.5px;'>$headerLabel</td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td style='padding:12px 16px 12px 16px;'>
            $rows
        </td>
    </tr>
</table>
<!--[if mso]>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td height="16" style="font-size:0; line-height:16px; mso-line-height-rule:exactly;">&#160;</td></tr></table>
<![endif]-->
"@
}

function Get-MailBodyHtmlHC {
    <#
    .SYNOPSIS
        The complete e-mail body, in the layout of the Permission matrix
        script: title, summary pills, system errors, a card per computer and
        a footer with the run times.

    .DESCRIPTION
        Table based with inline styles, so Outlook Classic (Word engine) and
        browsers show the same picture. Outlook gets a fixed width, browsers
        a fluid one capped at 900px.

    .PARAMETER Job
        Objects with ComputerName, Entries, Path (sort key), Description,
        Removed and Errors. Entries contain Name, Path and LinkPath; legacy
        single-path objects may supply these properties directly.

    .PARAMETER Body
        Optional HTML shown below the title. Used as supplied, without HTML
        encoding. Summary counts and task results are rendered separately,
        even when Body is empty.
    #>
    param (
        [Parameter(Mandatory)]
        [String]$ScriptName,
        [String]$Body,
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Job = @(),
        [Int]$Removed,
        [Int]$Errors,
        [String[]]$SystemError,
        [String]$LogFolderPath,
        [Boolean]$HasAttachments,
        [Parameter(Mandatory)]
        [DateTime]$ScriptStartTime,
        [DateTime]$ScriptEndTime = (Get-Date)
    )

    $theme = Get-MailThemeHC

    $banner = Build-MailSummaryBannerHC -Removed $Removed -Errors $Errors
    $systemErrorsBlock = Build-MailSystemErrorsBlockHC -Message $SystemError

    #region Computer cards, the ones with errors first
    $cards = $Job | Group-Object -Property ComputerName | Sort-Object -Property @{
        Expression = {
            if (($_.Group | Measure-Object -Property Errors -Sum).Sum) { 0 }
            elseif (($_.Group | Measure-Object -Property Removed -Sum).Sum) { 1 }
            else { 2 }
        }
    }, Name | ForEach-Object {
        Build-MailComputerCardHC -ComputerName $_.Name -Job $_.Group
    }
    #endregion

    #region Links
    $linkStyle = "color:$($theme.LinkColor); text-decoration:none; font-weight:600;"
    $links = @()

    if ($LogFolderPath) {
        $links += "<a href='$([System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $LogFolderPath)))' target='_blank' rel='noopener noreferrer' style='$linkStyle'>Open log folder</a>"
    }
    if ($HasAttachments) {
        $links += 'details in the attachments'
    }

    $linksBlock = if ($links) {
        "<p style='margin:0 0 14px 0; color:$($theme.TextMuted); font-size:12px; line-height:17px; mso-line-height-rule:exactly;'>$($links -join ' &middot; ')</p>"
    }
    #endregion

    #region Footer
    $footLabelStyle = "font-size:10px; font-weight:700; color:$($theme.TextLight); text-transform:uppercase; letter-spacing:0.5px;"
    $footValueStyle = "font-size:11px; color:$($theme.TextLight); font-family:$($theme.MonoStack);"
    $span = $ScriptEndTime - $ScriptStartTime

    $footer = @"
<table role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin:16px auto 0 auto;">
    <tr>
        <td style="padding:0 5px 0 0; $footLabelStyle">Started</td>
        <td style="padding:0 20px 0 0; $footValueStyle">$($ScriptStartTime.ToString('dd/MM/yyyy HH:mm'))</td>
        <td style="padding:0 5px 0 0; $footLabelStyle">Ended</td>
        <td style="padding:0 20px 0 0; $footValueStyle">$($ScriptEndTime.ToString('dd/MM/yyyy HH:mm'))</td>
        <td style="padding:0 5px 0 0; $footLabelStyle">Duration</td>
        <td style="padding:0; $footValueStyle">$('{0:00}:{1:00}:{2:00}' -f [int][Math]::Floor($span.TotalHours), $span.Minutes, $span.Seconds)</td>
    </tr>
</table>
<p style="margin:6px 0 0 0; text-align:center; $footValueStyle">$([System.Net.WebUtility]::HtmlEncode("$env:USERDNSDOMAIN\$env:USERNAME on $env:COMPUTERNAME, PowerShell $($PSVersionTable.PSVersion)"))</p>
"@
    #endregion

    $noJobs = if (-not $cards) {
        "<p style='margin:0; color:$($theme.TextLight); font-style:italic;'>No tasks were executed.</p>"
    }

    @"
<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if mso]>
<style type="text/css">
    v\:* { behavior: url(#default#VML); display:inline-block; }
    table { mso-table-lspace:0pt; mso-table-rspace:0pt; }
    td { mso-line-height-rule:exactly; }
</style>
<![endif]-->
<style type="text/css">
    body {
        font-family: $($theme.FontStack);
        font-size: 13px;
        color: $($theme.TextMain);
        background-color: $($theme.BgPage);
        margin: 0;
        padding: 0;
    }
    a { color: $($theme.LinkColor); text-decoration: none; }
    h1 { font-size: 22px; font-weight: 700; color: $($theme.TextMain); margin: 0 0 4px 0; }
    table { border-collapse: collapse; }
</style>
<!--[if !mso]><!-->
<style type="text/css">
    table.mail-root { max-width: 900px !important; }
</style>
<!--<![endif]-->
</head>
<body style="margin:0; padding:0;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="$($theme.BgPage)" style="border-collapse:collapse; background-color:$($theme.BgPage);">
    <tr>
        <td align="center" valign="top" bgcolor="$($theme.BgPage)" style="padding:20px; background-color:$($theme.BgPage);">
            <!--[if mso]>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="$($theme.BodyWidth)" align="center"><tr><td>
            <![endif]-->
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="mail-root" style="border-collapse:collapse; width:100%; margin:0 auto; text-align:left;">
                <tr><td style="padding:0 0 4px 0;"><h1>$([System.Net.WebUtility]::HtmlEncode($ScriptName))</h1></td></tr>
                <tr><td style="padding:0 0 16px 0; color:$($theme.TextMuted); font-size:13px; line-height:1.6;">$Body</td></tr>
                <tr><td style="padding:0;">$linksBlock</td></tr>
                <tr><td style="padding:0;">$banner</td></tr>
                <tr><td style="padding:0;">$systemErrorsBlock</td></tr>
                <tr><td style="padding:0;">$($cards -join '')$noJobs</td></tr>
                <tr><td style="padding:0;">$footer</td></tr>
            </table>
            <!--[if mso]>
            </td></tr></table>
            <![endif]-->
        </td>
    </tr>
</table>
</body>
</html>
"@
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
        Write log records to JSON or text files and return their paths.

    .DESCRIPTION
        JSON output keeps only DateTime and Message from each record;
        Message is converted to a string and all other properties are
        discarded. Text output formats all properties as a list.

        A failed export writes a warning instead of throwing. Only paths
        from completed exports are returned. The parent folder must exist.

    .PARAMETER DataToExport
        Records to write. For JSON, supply DateTime and Message properties.

    .PARAMETER PartialPath
        Output path without an extension, for example C:\Logs\Errors.

    .PARAMETER FileExtensions
        One or both of '.json' and '.txt'. Duplicate extensions are ignored.

    .PARAMETER Append
        Preserve existing content. For JSON, new records are placed before
        existing records and the file is rewritten. For text, new records
        are appended at the end. Without this switch, the file is replaced.

    .OUTPUTS
        String. Paths of successfully written log files.

    .EXAMPLE
        Out-LogFileHC -DataToExport ([pscustomobject]@{ DateTime = Get-Date; Message = 'Run completed' }) -PartialPath 'C:\Logs\Run' -FileExtensions '.json'

        Writes the record to C:\Logs\Run.json. C:\Logs must already exist.
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

        Supply paths to assemblies compatible with the installed PowerShell
        runtime. At least one To or Bcc recipient is required. A connection,
        authentication or sending failure throws an error.

    .PARAMETER MailKitAssemblyPath
        Path to MailKit.dll. Used if MailKit is not already loaded.

    .PARAMETER MimeKitAssemblyPath
        Path to MimeKit.dll. Used if MimeKit is not already loaded.

    .PARAMETER SmtpServerName
        SMTP server hostname or IP address.

    .PARAMETER SmtpPort
        SMTP port: 25, 465, 587 or 2525.

    .PARAMETER Body
        Email body as HTML.

    .PARAMETER Subject
        Complete subject line. This helper does not add counts or a prefix.

    .PARAMETER From
        Sender email address.

    .PARAMETER FromDisplayName
        Optional friendly name for the sender.

    .PARAMETER To
        Recipient addresses. To or Bcc must contain at least one address.

    .PARAMETER Bcc
        Blind-copy recipient addresses.

    .PARAMETER MaxAttachmentSize
        Combined source-file size limit in bytes. Defaults to 20 MB. At or
        above the limit, no attachments are added and a notice is appended
        to the body. This is not the size of the encoded email.

    .PARAMETER SmtpConnectionType
        MailKit SecureSocketOptions value. Defaults to None (no encryption).
        Use the mode required by your SMTP server, such as StartTls or
        SslOnConnect.

    .PARAMETER Priority
        Normal (default), Low or High. Sets the X-Priority header.

    .PARAMETER Attachments
        File paths to attach. Duplicates are removed. Missing files, folders
        and attachment failures produce warnings; sending can still proceed.

    .PARAMETER Credential
        Optional SMTP credential. If omitted, authentication is not attempted.
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
            <#
            .SYNOPSIS
                Check whether an assembly with the given name is already
                loaded in this process, regardless of its version.
            #>
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
            <#
            .SYNOPSIS
                Add readable files to a MIME body within the attachment limit.

            .DESCRIPTION
                Uses the enclosing function's MaxAttachmentSize. If the
                combined size reaches the limit, returns an object with
                AttachmentLimitExceededMessage and adds no attachments.
                Other attachment problems are warnings.
            #>
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
