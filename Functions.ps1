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


function Export-ExcelLogHC {
    <#
    .SYNOPSIS
        Exports log rows to numbered worksheets without exceeding Excel's row limit.

    .DESCRIPTION
        Reserves one row per worksheet for headers. The first worksheet keeps
        its name; subsequent worksheets and tables use _2, _3 and so on.

    .PARAMETER RowsPerSheet
        Maximum data rows per worksheet, excluding the header. A smaller value
        can be used to verify rollover without generating a million-row file.
    #>
    param (
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Rows,
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [ValidateSet('Overview', 'Errors')]
        [string]$WorksheetName,
        [ValidateRange(1, 1048575)]
        [int]$RowsPerSheet = 1048575
    )

    for ($offset = 0; $offset -lt $Rows.Count; $offset += $RowsPerSheet) {
        $sheetNumber = [int][Math]::Floor($offset / $RowsPerSheet) + 1
        $sheetName = if ($sheetNumber -eq 1) { $WorksheetName } else { "${WorksheetName}_$sheetNumber" }
        $end = [Math]::Min($offset + $RowsPerSheet, $Rows.Count)
        $exportParams = @{
            Path               = $Path
            WorksheetName      = $sheetName
            TableName          = $sheetName
            NoNumberConversion = '*'
            AutoSize           = $true
            FreezeTopRow       = $true
        }
        & {
            for ($rowIndex = $offset; $rowIndex -lt $end; $rowIndex++) {
                $Rows[$rowIndex]
            }
        } | Export-Excel @exportParams
    }
}

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
        ExcludeAttributes lists the protected attribute names.
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
    if ($Task.ExcludeAttributes) {
        $excluded += (($Task.ExcludeAttributes | ForEach-Object { $_.ToLower() } | Select-Object -Unique) -join '/') + ' items'
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

function Get-MailPathGroupHC {
    <#
    .SYNOPSIS
        Groups paths from one input task under shared parent folders.

    .DESCRIPTION
        Drive and UNC-share boundaries stay separate. Within each boundary,
        repeated child branches get their longest common parent. Remaining
        paths share a breadcrumb only when at least two distinct paths remain.
        This function inspects path strings only, without filesystem access.
    #>
    param (
        [Parameter(Mandatory)]
        [PSCustomObject[]]$Job
    )

    $commonParent = {
        param ([PSCustomObject[]]$Rows)
        $paths = @($Rows.Path | ForEach-Object { $_.Replace('/', '\').TrimEnd('\') } | Sort-Object -Unique)
        if ($paths.Count -lt 2) { return '' }
        $root = [IO.Path]::GetDirectoryName($paths[0])
        while ($root) {
            $prefix = $root.TrimEnd('\') + '\'
            if (-not @($paths | Where-Object { -not $_.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) }).Count) {
                return $root.TrimEnd('\')
            }
            $root = [IO.Path]::GetDirectoryName($root.TrimEnd('\'))
        }
        return ''
    }

    foreach ($volume in ($Job | Group-Object { [IO.Path]::GetPathRoot($_.Path.Replace('/', '\')) })) {
        $root = & $commonParent $volume.Group
        if (-not $root) {
            foreach ($row in $volume.Group) { [PSCustomObject]@{ Root = ''; Jobs = @($row) } }
            continue
        }
        $remaining = @()
        foreach ($branch in ($volume.Group | Group-Object { $_.Path.Replace('/', '\').Substring($root.Length + 1).Split('\')[0] })) {
            $branchRoot = & $commonParent $branch.Group
            if ($branchRoot) {
                [PSCustomObject]@{ Root = $branchRoot; Jobs = @($branch.Group) }
            }
            else { $remaining += $branch.Group }
        }
        if ($remaining.Count) {
            $remainingRoot = & $commonParent $remaining
            if ($remainingRoot) {
                [PSCustomObject]@{ Root = $remainingRoot; Jobs = $remaining }
            }
            else {
                foreach ($row in $remaining) { [PSCustomObject]@{ Root = ''; Jobs = @($row) } }
            }
        }
    }
}

function Build-MailJobRowHC {
    <#
    .SYNOPSIS
        One compact table row per task path, with removal and error counters.

    .PARAMETER Job
        Object containing Entries, Removed and Errors. Each entry
        contains Name, Path and LinkPath. A single-path job can supply those
        properties directly. Name is optional and appears above Path.
        Removed and Errors are counts. Text is HTML-encoded before rendering.
        Classic Outlook path labels are limited to 55 characters, preserving
        trailing components where possible. Browsers and link targets retain
        the full path, also supplied as the shortened label's tooltip.

    .PARAMETER RootPath
        Shared parent already displayed above at least two paths in this task.
        Renders relative child labels without changing links or job data.
    #>
    param (
        [Parameter(Mandatory)]
        [PSCustomObject]$Job,
        [String]$RootPath
    )

    $theme = Get-MailThemeHC

    $accent = if ($Job.Errors) {
        $theme.AccentError
    }
    else {
        $theme.TextLight
    }
    $rowBackground = if ($Job.Errors) { $theme.StatusError } else { $theme.BgWhite }
    $titleColor = if ($Job.Errors) { $theme.AccentError } else { $theme.TextMain }
    $detailColor = if ($Job.Errors) { $theme.AccentError } else { $theme.TextMuted }

    $entries = if ($Job.Entries) { $Job.Entries } else { @($Job) }
    $titleHtml = (@(foreach ($entry in $entries) {
        $href = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $entry.LinkPath))
        $path = [System.Net.WebUtility]::HtmlEncode($entry.Path)
        $displayPath = if ($RootPath) { $entry.Path.Substring($RootPath.Length + 1) } else { $entry.Path }
        $pathLabel = [System.Net.WebUtility]::HtmlEncode($displayPath)
        if ($displayPath.Length -gt 55) {
            $tail = $displayPath.TrimEnd([char[]]'\/')
            $tail = $tail.Substring([Math]::Max(0, $tail.Length - 52))
            $separatorIndex = $tail.IndexOfAny([char[]]'\/')
            $shortPath = if ($separatorIndex -ge 0) {
                '...\' + $tail.Substring($separatorIndex + 1)
            }
            else {
                '...' + $tail
            }
            $shortPath = [System.Net.WebUtility]::HtmlEncode($shortPath)
            $pathLabel = "<!--[if mso]><span title='$path'>$shortPath</span><![endif]--><!--[if !mso]><!-->$pathLabel<!--<![endif]-->"
        }

        if ($RootPath) {
            $label = if ($entry.Name) { [System.Net.WebUtility]::HtmlEncode($entry.Name) } else { $pathLabel }
            $suffix = if ($entry.Name) { " <span style='font-weight:400;'>($pathLabel)</span>" } else { '' }
            "<p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; font-family:$($theme.MonoStack); font-weight:400; font-size:11px; color:$titleColor; line-height:16px; mso-line-height-rule:exactly; overflow-wrap:anywhere; word-break:break-all;'><a href='$href' title='$path' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$titleColor;'>$label</a>$suffix</p>"
        }
        elseif ($entry.Name) {
            "<p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; font-weight:700; color:$titleColor; font-size:13px; line-height:16px; mso-line-height-rule:exactly;'><a href='$href' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$titleColor;'>$([System.Net.WebUtility]::HtmlEncode($entry.Name))</a></p>" +
            "<p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; font-family:$($theme.MonoStack); font-size:11px; color:$detailColor; line-height:14px; mso-line-height-rule:exactly; overflow-wrap:anywhere; word-break:break-all;'>$pathLabel</p>"
        }
        else {
            "<p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; font-family:$($theme.MonoStack); font-weight:700; font-size:12px; color:$titleColor; line-height:16px; mso-line-height-rule:exactly; overflow-wrap:anywhere; word-break:break-all;'><a href='$href' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$titleColor;'>$pathLabel</a></p>"
        }
    })) -join ''

    @"
    <tr class='path-row'>
        <td valign='middle' bgcolor='$rowBackground' style='vertical-align:middle; padding:8px;$(if ($RootPath) { ' padding-left:16px;' }) background-color:$rowBackground; color:$titleColor; border-bottom:1px solid $($theme.BorderLight);'>
            $titleHtml
        </td>
        <td class='removed-count' valign='middle' align='right' width='64' bgcolor='$rowBackground' style='vertical-align:middle; padding:8px; background-color:$rowBackground; border-bottom:1px solid $($theme.BorderLight); color:$detailColor; font-size:12px; line-height:16px; mso-line-height-rule:exactly; text-align:right;'>$($Job.Removed)</td>
        <td class='error-count' valign='middle' align='right' width='48' bgcolor='$rowBackground' style='vertical-align:middle; padding:8px; background-color:$rowBackground; border-bottom:1px solid $($theme.BorderLight); color:$accent; font-weight:700; font-size:12px; line-height:16px; mso-line-height-rule:exactly; text-align:right;'>$($Job.Errors)</td>
    </tr>
"@
}

function Build-MailComputerCardHC {
    <#
    .SYNOPSIS
        A card per executing computer, with task descriptions and path counters.

    .DESCRIPTION
        The header is red when a job failed, green when something was
        removed and grey when nothing was removed. Outlook cannot render the
        gradient, so it gets the average color of the two gradient stops.
        Rows sharing a TaskIndex have one description above their table.
        Shared subfolders group at least two distinct paths within that task;
        singletons stay ungrouped. Tasks, path groups and rows with errors
        appear first, then sort by Path, without prioritizing removals.
        Error rows have red text on a pale-red background across all cells.
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

    $taskGroups = $Job | Group-Object -Property {
        if ($null -ne $_.TaskIndex) { "Task:$($_.TaskIndex)" }
        else { "Row:$([array]::IndexOf($Job, $_))" }
    }
    $rows = ($taskGroups | Sort-Object -Property @{
            Expression = { if (($_.Group | Measure-Object -Property Errors -Sum).Sum) { 0 } else { 1 } }
        }, @{
            Expression = { $_.Group.Path | Sort-Object | Select-Object -First 1 }
        }, Name | ForEach-Object {
            $description = [System.Net.WebUtility]::HtmlEncode(($_.Group.Description | Select-Object -Unique) -join '; ')
            $pathGroups = Get-MailPathGroupHC -Job $_.Group | Sort-Object -Property @{
                Expression = { if (($_.Jobs | Measure-Object -Property Errors -Sum).Sum) { 0 } else { 1 } }
            }, @{ Expression = { $_.Jobs.Path | Sort-Object | Select-Object -First 1 } }
            $pathRows = (@(foreach ($pathGroup in $pathGroups) {
                if ($pathGroup.Root) {
                    $root = [System.Net.WebUtility]::HtmlEncode($pathGroup.Root + '\')
                    "<tr class='root-breadcrumb'><td colspan='3' bgcolor='#f3f4f6' style='padding:8px; background-color:#f3f4f6; border-bottom:1px solid $($theme.BorderMain); color:$($theme.TextMuted); font-family:$($theme.MonoStack); font-size:12px; font-weight:700; line-height:16px; mso-line-height-rule:exactly; overflow-wrap:anywhere; word-break:break-all;'><p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; font-family:$($theme.MonoStack); font-size:12px; font-weight:700; line-height:16px; mso-line-height-rule:exactly;'><strong style='font-weight:700;'>$root</strong></p></td></tr>"
                }
                foreach ($row in ($pathGroup.Jobs | Sort-Object -Property @{
                    Expression = { if ($_.Errors) { 0 } else { 1 } }
                }, Path)) {
                    Build-MailJobRowHC -Job $row -RootPath $pathGroup.Root
                }
            })) -join ''
            @"
<!--[if mso]>
<table class='task-description' role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='border-collapse:collapse; width:100%;'>
    <tr><td style='text-align:left; padding:8px 8px 6px; font-size:12px; font-weight:400; color:$($theme.TextMain); line-height:17px; mso-line-height-rule:exactly;'><p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; line-height:17px; mso-line-height-rule:exactly;'>$description</p></td></tr>
</table>
<![endif]-->
<table class="task-table" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; width:100%; table-layout:fixed; margin:0 0 16px 0;">
    <!--[if !mso]><!-->
    <caption style='text-align:left; padding:8px 8px 6px; font-size:12px; font-weight:400; color:$($theme.TextMain); line-height:17px; mso-line-height-rule:exactly;'>$description</caption>
    <!--<![endif]-->
    <tr>
        <th scope='col' align='left' style='padding:6px 8px; border-bottom:1px solid $($theme.BorderMain); color:$($theme.TextLight); font-size:11px;'><p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; font-size:11px; font-weight:700; line-height:15px; mso-line-height-rule:exactly;'>Path</p></th>
        <th scope='col' align='right' width='64' style='padding:6px 8px; border-bottom:1px solid $($theme.BorderMain); color:$($theme.TextLight); font-size:11px;'><p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; font-size:11px; font-weight:700; line-height:15px; mso-line-height-rule:exactly;'>Removed</p></th>
        <th scope='col' align='right' width='48' style='padding:6px 8px; border-bottom:1px solid $($theme.BorderMain); color:$($theme.TextLight); font-size:11px;'><p style='margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0; font-size:11px; font-weight:700; line-height:15px; mso-line-height-rule:exactly;'>Errors</p></th>
    </tr>
    $pathRows
</table>
"@
        }) -join ''

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
                        <p style='margin:0; font-size:16px; font-weight:700; color:#ffffff; line-height:20px; mso-line-height-rule:exactly; overflow-wrap:anywhere; word-break:break-word;'>$([System.Net.WebUtility]::HtmlEncode($ComputerName))</p>
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
        Objects with TaskIndex, ComputerName, Entries, Path (sort key), Description,
        Removed and Errors. Entries contain Name, Path and LinkPath; legacy
        single-path objects may supply these properties directly.

    .PARAMETER Body
        Optional HTML shown below the title. Used as supplied, without HTML
        encoding. Summary counts and task results are rendered separately,
        even when Body is empty.

    .PARAMETER BrowserViewFilePath
        Saved mail HTML path. Shows the browser-view link in classic Outlook
        only; omit when no HTML copy is available.
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
        [DateTime]$ScriptEndTime = (Get-Date),
        [String]$BrowserViewFilePath
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
    $browserViewRow = ''

    if (-not [string]::IsNullOrWhiteSpace($BrowserViewFilePath)) {
        $browserUrl = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $BrowserViewFilePath))
        $browserTitle = [System.Net.WebUtility]::HtmlEncode($BrowserViewFilePath)
        $browserViewRow = @"
<!--[if mso]>
<tr><td style='padding:0 0 8px 0; color:$($theme.TextMuted); font-size:12px;'><p style='margin:0; mso-line-height-rule:exactly; line-height:17px;'>If this mail is not visible, please <a href='$browserUrl' title="$browserTitle" target='_blank' rel='noopener noreferrer' style='$linkStyle'>click here to view it in the browser</a>.</p></td></tr>
<![endif]-->
"@
    }

    if ($LogFolderPath) {
        $folderPath = [System.Net.WebUtility]::HtmlEncode($LogFolderPath)
        $folderUrl = [System.Net.WebUtility]::HtmlEncode(([uri]$LogFolderPath).AbsoluteUri)
        $links += @"
<!--[if mso]><a href='$folderPath' title='$folderPath' style='$linkStyle'>Open log folder</a><![endif]-->
<!--[if !mso]><!--><a href='$folderUrl' title='$folderPath' style='$linkStyle'>Open log folder</a><!--<![endif]-->
"@
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
<table class="mail-footer" role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin:16px auto 0 auto;">
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
    @media screen and (max-width: 480px) {
        table.mail-footer tr { display:grid; grid-template-columns:auto 1fr; }
        table.mail-footer td { padding:2px 5px !important; }
    }
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
                $browserViewRow
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
