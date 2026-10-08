#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.2.0' }

BeforeAll {
    . (Join-Path -Path (Split-Path $PSScriptRoot) -ChildPath 'Functions.ps1')
}
Describe 'Export-ExcelLogHC' {
    It 'splits <Count> rows without losing or duplicating data' -ForEach @(
        @{ Count = 0; ExpectedSheets = 0 }
        @{ Count = 1; ExpectedSheets = 1 }
        @{ Count = 2; ExpectedSheets = 1 }
        @{ Count = 3; ExpectedSheets = 2 }
        @{ Count = 4; ExpectedSheets = 2 }
        @{ Count = 5; ExpectedSheets = 3 }
    ) {
        $path = Join-Path $TestDrive "rows-$Count.xlsx"
        $rows = @(for ($rowIndex = 0; $rowIndex -lt $Count; $rowIndex++) {
            [pscustomobject]@{ Id = $rowIndex; Code = '001'; Error = 'Example error' }
        })
        Export-ExcelLogHC -Rows $rows -Path $path -WorksheetName Overview -RowsPerSheet 2
        if ($Count -eq 0) {
            Test-Path -LiteralPath $path | Should-BeFalse
            return
        }
        $sheets = @(Get-ExcelSheetInfo -Path $path)
        $sheets | Should-BeCollection -Count $ExpectedSheets
        $actualRows = @(foreach ($sheet in $sheets) {
            Import-Excel -Path $path -WorksheetName $sheet.Name
        })
        $actualRows.Id | Should-BeCollection @($rows.Id)
        $package = Open-ExcelPackage -Path $path
        try {
            for ($sheetIndex = 0; $sheetIndex -lt $ExpectedSheets; $sheetIndex++) {
                $expectedName = if ($sheetIndex -eq 0) { 'Overview' } else { 'Overview_{0}' -f ($sheetIndex + 1) }
                $sheets[$sheetIndex].Name | Should-Be $expectedName
                $sheet = $package.Workbook.Worksheets[$expectedName]
                $sheet.Dimension.End.Row | Should-Be ([Math]::Min(2, $Count - 2 * $sheetIndex) + 1)
                $sheet.Cells[1, 1].Text | Should-Be 'Id'
                $sheet.Cells[2, 2].Text | Should-Be '001'
                $sheet.Tables[0].Name | Should-Be $expectedName
                $pane = $sheet.WorksheetXml.SelectSingleNode("//*[local-name()='pane']")
                $pane.topLeftCell | Should-Be 'A2'
                $pane.state | Should-Be 'frozen'
            }
        }
        finally { $package.Dispose() }
    }
    It 'keeps Overview sheets when Errors also rolls over in the same workbook' {
        $path = Join-Path $TestDrive 'combined.xlsx'
        $rows = @([pscustomobject]@{ Error = 'First' }, [pscustomobject]@{ Error = 'Second' })
        foreach ($name in @('Overview', 'Errors')) {
            Export-ExcelLogHC -Rows $rows -Path $path -WorksheetName $name -RowsPerSheet 1
        }
        @(Get-ExcelSheetInfo -Path $path).Name | Should-BeCollection @('Overview', 'Overview_2', 'Errors', 'Errors_2')
        foreach ($name in @('Overview', 'Errors')) {
            (Import-Excel -Path $path -WorksheetName $name).Error | Should-Be 'First'
            (Import-Excel -Path $path -WorksheetName "${name}_2").Error | Should-Be 'Second'
        }
    }
    It 'reserves one header row within the Excel maximum by default' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path (Split-Path $PSScriptRoot) 'Functions.ps1'), [ref]$null, [ref]$null
        )
        $function = $ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Export-ExcelLogHC'
        }, $true)
        $parameter = $function.Body.ParamBlock.Parameters |
            Where-Object { $_.Name.VariablePath.UserPath -eq 'RowsPerSheet' }
        $parameter.DefaultValue.SafeGetValue() | Should-Be 1048575
        { Export-ExcelLogHC -Rows @() -Path 'unused.xlsx' -WorksheetName Overview -RowsPerSheet 1048576 } | Should-Throw
        { Export-ExcelLogHC -Rows @() -Path 'unused.xlsx' -WorksheetName Overview -RowsPerSheet 0 } | Should-Throw
    }
}
Describe 'ConvertTo-FileUrlHC' {
    It 'converts <Path>' -ForEach @(
        @{ Path = '\\server\share\my folder'; Expected = 'file:////server/share/my%20folder' }
        @{ Path = 'C:\Temp'; Expected = 'file://C:/Temp' }
        @{ Path = ''; Expected = '' }
    ) {
        ConvertTo-FileUrlHC -Path $Path | Should-Be $Expected
    }
}
Describe 'New-PillHtmlHC' {
    It 'renders a browser span and an Outlook VML shape' {
        $actual = New-PillHtmlHC -Text 'Error' -Bg '#dc2626'

        $actual | Should-BeLikeString '*<!--`[if mso`]>*<v:roundrect*fillcolor="#dc2626"*>ERROR</center>*<!`[endif`]-->*'
        $actual | Should-BeLikeString '*<!--`[if !mso`]><!-->*background-color:#dc2626*>Error</span>*'
    }
    It 'renders nothing without text' {
        New-PillHtmlHC -Text '' -Bg '#dc2626' | Should-BeEmptyString
    }
}
Describe 'Get-TaskDescriptionHC' {
    It 'shows active attribute exclusions' {
        Get-TaskDescriptionHC -Task ([pscustomobject]@{
            Type = 'RemoveFilesInFolder'
            ExcludeAttributes = @('Hidden', 'System')
        }) | Should-Be 'Remove all files, excluding hidden/system items'
    }
    It '<Expected>' -ForEach @(
        @{
            Task     = @{ Type = 'RemoveFile'; OlderThan = @{ Quantity = 7; Unit = 'Day'; BasedOn = 'LastWriteTime' } }
            Expected = 'Remove file older than 7 days (last write time)'
        }
        @{
            Task     = @{ Type = 'RemoveFile'; OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'CreationTime' } }
            Expected = 'Remove file'
        }
        @{
            Task     = @{ Type = 'RemoveFilesInFolder'; Recurse = $true; OlderThan = @{ Quantity = 1; Unit = 'Month'; BasedOn = 'CreationTime' } }
            Expected = 'Remove files older than 1 month (creation time), including subfolders'
        }
        @{
            Task     = @{ Type = 'RemoveFilesInFolder'; Recurse = $false; OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'CreationTime' } }
            Expected = 'Remove all files'
        }
        @{
            Task     = @{ Type = 'RemoveFilesInFolder'; Recurse = $true; OlderThan = @{ Quantity = 3; Unit = 'Year'; BasedOn = 'CreationTime' }; ExcludeFolders = @('a', 'b'); ExcludeFiles = @('c') }
            Expected = 'Remove files older than 3 years (creation time), including subfolders, excluding 2 folders and 1 file'
        }
        @{
            Task     = @{ Type = 'RemoveEmptyFolders'; ExcludeFolders = @('a'); ExcludeFiles = @('c') }
            Expected = 'Remove empty folders, excluding 1 folder'
        }
    ) {
        Get-TaskDescriptionHC -Task ([PSCustomObject]$Task) | Should-Be $Expected
    }
}
Describe 'Get-MailBodyHtmlHC' {
    BeforeAll {
        $testTheme = Get-MailThemeHC

        $testJobs = @(
            [PSCustomObject]@{ ComputerName = 'PC-OK'; Name = 'Logs'; Path = 'D:\Logs'; LinkPath = '\\PC-OK\D$\Logs'; Description = 'Remove all files'; Removed = 3; Errors = 0 }
            [PSCustomObject]@{ ComputerName = 'PC-OK'; Name = $null; Path = 'D:\Idle'; LinkPath = '\\PC-OK\D$\Idle'; Description = 'Remove empty folders'; Removed = 0; Errors = 0 }
            [PSCustomObject]@{ ComputerName = 'PC-ERR'; Name = $null; Path = 'E:\<Data>'; LinkPath = '\\PC-ERR\E$\<Data>'; Description = 'Remove all files'; Removed = 1; Errors = 2 }
            [PSCustomObject]@{ ComputerName = 'PC-IDLE'; Name = $null; Path = 'F:\Empty'; LinkPath = '\\PC-IDLE\F$\Empty'; Description = 'Remove all files'; Removed = 0; Errors = 0 }
        )

        $testParams = @{
            ScriptName      = 'Remove <test>'
            Body            = '<p>Custom body</p>'
            Job             = $testJobs
            Removed         = 4
            Errors          = 3
            SystemError     = @('Oops & co')
            LogFolderPath   = 'C:\Log folder'
            HasAttachments  = $true
            ScriptStartTime = Get-Date '2026-10-02 07:00'
            ScriptEndTime   = Get-Date '2026-10-02 08:30:15'
        }

        $actual = Get-MailBodyHtmlHC @testParams
    }
    It 'shows the encoded script name as title and the user body as is' {
        $actual | Should-BeLikeString '*<h1>Remove &lt;test&gt;</h1>*'
        $actual | Should-BeLikeString '*<p>Custom body</p>*'
    }
    It 'has a fixed width for Outlook' {
        $actual | Should-BeLikeString "*<!--``[if mso``]>*width=`"$($testTheme.BodyWidth)`"*"
    }
    It 'shows the totals as pills' {
        $actual | Should-BeLikeString '*>4 Removed</span>*'
        $actual | Should-BeLikeString '*>3 Errors</span>*'
    }
    It 'shows the encoded system errors' {
        $actual | Should-BeLikeString '*System Errors (1)*Oops &amp; co*'
    }
    It 'links the log folder and mentions the attachments' {
        $actual | Should-BeLikeString "*href='C:\Log folder'*Open log folder*details in the attachments*"
    }
    It 'uses client-specific folder links for <Path>' -ForEach @(
        @{ Path = '\\server\Logs\Team A & B'; Url = 'file://server/Logs/Team%20A%20&amp;%20B' }
        @{ Path = 'C:\Log folder\Run #1'; Url = 'file:///C:/Log%20folder/Run%20%231' }
    ) {
        $testNewParams = $testParams.Clone()
        $testNewParams.LogFolderPath = $Path

        $html = Get-MailBodyHtmlHC @testNewParams

        $outlookLink = [regex]::Match($html, '<!--\[if mso\]>(<a [^>]+>Open log folder</a>)<!\[endif\]-->').Groups[1].Value
        $browserLink = [regex]::Match($html, '<!--\[if !mso\]><!-->(<a [^>]+>Open log folder</a>)<!--<!\[endif\]-->').Groups[1].Value
        $outlookLink.Contains("href='$([System.Net.WebUtility]::HtmlEncode($Path))'") | Should-BeTrue
        $outlookLink | Should-NotBeLikeString '*target=*'
        $outlookLink | Should-NotBeLikeString '*file://*'
        $browserLink.Contains("href='$Url'") | Should-BeTrue
    }
    It 'shows an encoded browser-view link only inside an Outlook conditional row' {
        $testNewParams = $testParams.Clone()
        $testNewParams.BrowserViewFilePath = '\\server\Logs\Team A & B - Mail.html'

        $html = Get-MailBodyHtmlHC @testNewParams

        $html | Should-MatchString '(?s)<!--\[if mso\]>\s*<tr><td[^>]*><p[^>]*>If this mail is not visible, please <a[^>]+>click here to view it in the browser</a>\.</p></td></tr>\s*<!\[endif\]-->'
        $html | Should-BeLikeString "*href='file:////server/Logs/Team%20A%20&amp;%20B%20-%20Mail.html'*"
        $html | Should-BeLikeString '*title="\\server\Logs\Team A &amp; B - Mail.html"*'
    }
    It 'omits the browser-view link without a saved HTML path' {
        $actual | Should-NotBeLikeString '*If this mail is not visible*'
    }
    It 'shows a card per computer, the one with errors first and the idle one last' {
        $errIndex = $actual.IndexOf('>PC-ERR</p>')
        $okIndex = $actual.IndexOf('>PC-OK</p>')
        $idleIndex = $actual.IndexOf('>PC-IDLE</p>')

        $errIndex | Should-BeGreaterThan 0
        $errIndex | Should-BeLessThan $okIndex
        $okIndex | Should-BeLessThan $idleIndex
    }
    It 'colors the card headers by result' {
        $actual | Should-BeLikeString "*linear-gradient(135deg, $($testTheme.GradError[0])*>PC-ERR</p>*"
        $actual | Should-BeLikeString "*linear-gradient(135deg, $($testTheme.GradSuccess[0])*>PC-OK</p>*"
        $actual | Should-BeLikeString "*linear-gradient(135deg, $($testTheme.GradIdle[0])*>PC-IDLE</p>*"
    }
    It 'shows the totals in the card header' {
        $actual | Should-BeLikeString '*>PC-ERR</p>*1&nbsp;removed &middot; 2&nbsp;errors*'
        $actual | Should-BeLikeString '*>PC-OK</p>*2 paths*3&nbsp;removed*'
    }
    It 'shows the task description above path rows with separate counters' {
        $actual | Should-BeLikeString '*Remove all files</caption>*'
        $actual | Should-BeLikeString "*href='file:////PC-OK/D$/Logs'*>Logs</a>*D:\Logs*class='removed-count'*>3</td>*"
        $actual | Should-BeLikeString "*E:\&lt;Data&gt;*class='removed-count'*>1</td>*class='error-count'*>2</td>*"
    }
    It 'highlights error counters without individual row cards or badges' {
        $actual | Should-BeLikeString "*class='error-count'*color:$($testTheme.AccentError)*>2</td>*"
        $pathRows = [regex]::Matches($actual, "(?s)<tr class='path-row'>.*?</tr>").Value -join ''
        $pathRows | Should-NotBeLikeString '*border-left:3px solid*'
        $actual | Should-NotBeLikeString '*>Error</span>*'
    }
    It 'renders the description once per task and keeps each path counter' {
        $html = Build-MailComputerCardHC -ComputerName 'PC1' -Job @(
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\First'; LinkPath = 'C:\First'; Description = 'Remove files & folders'; Removed = 5; Errors = 0 }
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\Second'; LinkPath = 'C:\Second'; Description = 'Remove files & folders'; Removed = 0; Errors = 2 }
            [pscustomobject]@{ TaskIndex = 1; Path = 'C:\First'; LinkPath = 'C:\First'; Description = 'Remove files & folders'; Removed = 0; Errors = 0 }
        )
        $browserHtml = [regex]::Replace($html, '(?s)<!--\[if mso\]>.*?<!\[endif\]-->', '')
        $outlookHtml = [regex]::Replace($html, '(?s)<!--\[if !mso\]><!-->.*?<!--<!\[endif\]-->', '')
        [regex]::Matches($browserHtml, 'Remove files &amp; folders').Count | Should-Be 2
        [regex]::Matches($outlookHtml, 'Remove files &amp; folders').Count | Should-Be 2
        $outlookHtml | Should-NotBeLikeString '*<caption*'
        $outlookHtml | Should-BeLikeString '*<table class=''task-description''*>Remove files &amp; folders</p></td></tr>*</table>*class="task-table"*'
        $browserHtml | Should-NotBeLikeString '*class=''task-description''*'
        foreach ($taskTable in [regex]::Matches($outlookHtml, '(?s)<table class="task-table".*?</table>')) {
            $table = ([xml]$taskTable.Value).table
            $table.tr[0].th | Should-BeCollection -Count 3
            $table.tr[0].th[1].width | Should-Be '64'
            $table.tr[0].th[2].width | Should-Be '48'
            $table.tr[0].th[1].align | Should-Be 'right'
            $table.tr[0].th[2].align | Should-Be 'right'
        }
        [regex]::Matches($html, 'class="task-table"').Count | Should-Be 2
        $pathRows = [regex]::Matches($html, "(?s)<tr class='path-row'>.*?</tr>").Value
        $pathRows | Should-BeCollection -Count 3
        $pathRows[0] | Should-BeLikeString "*C:\Second*class='removed-count'*>0</td>*class='error-count'*>2</td>*"
        $pathRows[1] | Should-BeLikeString "*C:\First*class='removed-count'*>5</td>*class='error-count'*>0</td>*"
    }
    It 'colors every cell and path label red only when the row has errors (<Name>)' -ForEach @(
        @{ Name = $null }
        @{ Name = 'Named & failed folder' }
    ) {
        foreach ($errorCount in @(0, 2)) {
            $html = Build-MailJobRowHC -Job ([pscustomobject]@{
                Path = 'C:\Failed'; LinkPath = 'C:\Failed'; Name = $Name; Removed = 3; Errors = $errorCount
            })
            $row = ([xml]"<table>$html</table>").table.tr
            $row.td | Should-BeCollection -Count 3
            foreach ($cell in $row.td) {
                if ($errorCount) {
                    $cell.bgcolor | Should-Be $testTheme.StatusError
                    $cell.style | Should-BeLikeString "*background-color:$($testTheme.StatusError)*color:$($testTheme.AccentError)*"
                }
                else {
                    $cell.bgcolor | Should-Be $testTheme.BgWhite
                    $cell.style | Should-NotBeLikeString "*color:$($testTheme.AccentError)*"
                }
            }
            foreach ($label in $row.SelectNodes('.//p | .//a')) {
                if ($errorCount) { $label.style | Should-BeLikeString "*color:$($testTheme.AccentError)*" }
                else { $label.style | Should-NotBeLikeString "*color:$($testTheme.AccentError)*" }
            }
        }
    }
    It 'uses zero-margin paragraphs in vertically centered path cells (<Name>, <RootPath>)' -ForEach @(
        @{ Name = $null; RootPath = $null }
        @{ Name = 'Named path'; RootPath = $null }
        @{ Name = $null; RootPath = 'C:\Root' }
        @{ Name = 'Named path'; RootPath = 'C:\Root' }
    ) {
        $html = Build-MailJobRowHC -Job ([pscustomobject]@{
            Path = 'C:\Root\File.log'; LinkPath = 'C:\Root\File.log'; Name = $Name; Removed = 1; Errors = 0
        }) -RootPath $RootPath
        $row = ([xml]"<table>$html</table>").table.tr
        $row.td[0].valign | Should-Be 'middle'
        $row.td[0].style | Should-BeLikeString '*vertical-align:middle*'
        $row.SelectNodes('.//div').Count | Should-Be 0
        $paragraphs = $row.td[0].SelectNodes('./p')
        $paragraphs.Count | Should-BeGreaterThan 0
        foreach ($paragraph in $paragraphs) {
            $paragraph.style | Should-BeLikeString '*margin:0; mso-margin-top-alt:0; mso-margin-bottom-alt:0;*'
            $paragraph.style | Should-BeLikeString '*mso-line-height-rule:exactly*'
        }
    }
    It 'sorts errors first and then by full Path without prioritizing removals' {
        $html = Build-MailComputerCardHC -ComputerName 'PC1' -Job @(
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\Zulu'; LinkPath = 'C:\Zulu'; Description = 'Later task'; Removed = 0; Errors = 1 }
            [pscustomobject]@{ TaskIndex = 1; Path = 'C:\Charlie'; LinkPath = 'C:\Charlie'; Description = 'Earlier task'; Removed = 0; Errors = 2 }
            [pscustomobject]@{ TaskIndex = 1; Path = 'C:\Delta'; LinkPath = 'C:\Delta'; Description = 'Earlier task'; Removed = 0; Errors = 1 }
            [pscustomobject]@{ TaskIndex = 1; Path = 'C:\Bravo'; LinkPath = 'C:\Bravo'; Description = 'Earlier task'; Removed = 5; Errors = 0 }
            [pscustomobject]@{ TaskIndex = 1; Path = 'C:\Alpha'; LinkPath = 'C:\Alpha'; Description = 'Earlier task'; Removed = 0; Errors = 0 }
        )
        $pathRows = [regex]::Matches($html, "(?s)<tr class='path-row'>.*?</tr>").Value
        $pathRows | Should-BeCollection -Count 5
        $pathRows[0] | Should-BeLikeString '*C:\Charlie*'
        $pathRows[1] | Should-BeLikeString '*C:\Delta*'
        $pathRows[2] | Should-BeLikeString '*C:\Alpha*'
        $pathRows[3] | Should-BeLikeString '*C:\Bravo*'
        $pathRows[4] | Should-BeLikeString '*C:\Zulu*'
    }
    It 'shows the run times in the footer' {
        $actual | Should-BeLikeString '*Started*02/10/2026 07:00*Ended*02/10/2026 08:30*Duration*01:30:15*'
    }
    It 'shows a message when there are no jobs' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Job = @()

        Get-MailBodyHtmlHC @testNewParams | Should-BeLikeString '*No tasks were executed.*'
    }
    It 'shows a grey removed pill and no error pill when nothing happened' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Removed = 0
        $testNewParams.Errors = 0

        $result = Get-MailBodyHtmlHC @testNewParams

        $result | Should-BeLikeString "*background-color:$($testTheme.AccentIdle)*>0 Removed</span>*"
        $result | Should-NotBeLikeString '*Errors</span>*'
    }
}
Describe 'Get-MailPathGroupHC' {
    It 'renders shared subfolders only inside the same task and keeps singletons unchanged' {
        $jobs = @(
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\Share\BE\a.txt'; LinkPath = '\\PC1\C$\Share\BE\a.txt'; Name = 'First & named'; Description = 'Rule'; Removed = 2; Errors = 0 }
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\Share\BE\b.txt'; LinkPath = '\\PC1\C$\Share\BE\b.txt'; Description = 'Rule'; Removed = 0; Errors = 1 }
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\Share\Other\only.txt'; LinkPath = '\\PC1\C$\Share\Other\only.txt'; Description = 'Rule'; Removed = 0; Errors = 0 }
            [pscustomobject]@{ TaskIndex = 1; Path = 'C:\Share\BE\c.txt'; LinkPath = '\\PC1\C$\Share\BE\c.txt'; Description = 'Rule'; Removed = 0; Errors = 0 }
        )
        $html = Build-MailComputerCardHC -ComputerName 'PC1' -Job $jobs
        [regex]::Matches($html, "class='root-breadcrumb'").Count | Should-Be 1
        $html | Should-BeLikeString '*C:\Share\BE\</td>*'
        $rows = [regex]::Matches($html, "(?s)<tr class='path-row'>.*?</tr>").Value
        $rows | Should-BeCollection -Count 4
        $rows[0] | Should-BeLikeString "*title='C:\Share\BE\b.txt'*>b.txt</a>*"
        $rows[0] | Should-BeLikeString '*background-color:#fee2e2*'
        $rows[1] | Should-BeLikeString "*href='file:////PC1/C$/Share/BE/a.txt'*title='C:\Share\BE\a.txt'*>First &amp; named</a>*`(a.txt`)*"
        $rows[1] | Should-NotBeLikeString '*<br>*'
        $rows[2] | Should-BeLikeString '*>C:\Share\Other\only.txt</a>*'
        $rows[3] | Should-BeLikeString '*>C:\Share\BE\c.txt</a>*'
        $jobs[0].Path | Should-Be 'C:\Share\BE\a.txt'
        $html | Should-BeLikeString '*2&nbsp;removed &middot; 1&nbsp;error*'
    }
    It 'emphasizes shared roots while keeping children and both client descriptions regular' {
        $jobs = @(
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\Root\First'; LinkPath = 'C:\Root\First'; Name = 'Named child'; Description = 'Cleanup rule'; Removed = 2; Errors = 0 }
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\Root\Second'; LinkPath = 'C:\Root\Second'; Description = 'Cleanup rule'; Removed = 0; Errors = 1 }
        )
        $html = Build-MailComputerCardHC -ComputerName 'PC1' -Job $jobs
        $root = [xml][regex]::Match($html, "(?s)<tr class='root-breadcrumb'>.*?</tr>").Value
        $root.tr.td.style | Should-BeLikeString '*font-weight:700;*'
        foreach ($pathRow in [regex]::Matches($html, "(?s)<tr class='path-row'>.*?</tr>")) {
            $row = [xml]$pathRow.Value
            $row.tr.td[0].p.style | Should-BeLikeString '*font-weight:400;*'
        }
        $caption = [xml][regex]::Match($html, '(?s)<caption.*?</caption>').Value
        $caption.caption.style | Should-BeLikeString '*font-weight:400;*'
        $description = [xml][regex]::Match($html, "(?s)<table class='task-description'.*?</table>").Value
        $description.table.tr.td.style | Should-BeLikeString '*font-weight:400;*'
    }
    It 'encodes shared roots and relative labels without changing UNC destinations' {
        $jobs = @('a & b.txt', "c 'd'.txt") | ForEach-Object {
            [pscustomobject]@{ TaskIndex = 0; Path = "\\server\share\A & B\$_"; LinkPath = "\\server\share\A & B\$_"; Description = 'Files'; Removed = 0; Errors = 0 }
        }
        $html = Build-MailComputerCardHC -ComputerName 'PC1' -Job $jobs
        $html | Should-BeLikeString '*\\server\share\A &amp; B\</td>*'
        $html | Should-BeLikeString "*href='file:////server/share/A%20&amp;%20B/a%20&amp;%20b.txt'*"
        $html | Should-BeLikeString '*>a &amp; b.txt</a>*'
        $html | Should-BeLikeString '*>c &#39;d&#39;.txt</a>*'
    }
    It 'groups repeated branches while leaving a lone unrelated path alone' {
        $jobs = @('C:\Share\BE\a.txt', 'C:\Share\BE\b.txt', 'C:\Share\NL\a.txt', 'C:\Share\NL\b.txt', 'C:\Share\Other\only.txt') | ForEach-Object { [pscustomobject]@{ Path = $_ } }
        $groups = @(Get-MailPathGroupHC -Job $jobs)
        $groups | Should-BeCollection -Count 3
        ($groups | Where-Object Root -EQ 'C:\Share\BE').Jobs | Should-BeCollection -Count 2
        ($groups | Where-Object Root -EQ 'C:\Share\NL').Jobs | Should-BeCollection -Count 2
        ($groups | Where-Object Root -EQ '').Jobs.Path | Should-Be 'C:\Share\Other\only.txt'
    }
    It 'keeps single paths and duplicate paths ungrouped' {
        foreach ($paths in @(@('C:\Share\only.txt'), @('C:\Share\only.txt', 'c:\share\ONLY.txt'))) {
            $jobs = @($paths | ForEach-Object { [pscustomobject]@{ Path = $_ } })
            $groups = @(Get-MailPathGroupHC -Job $jobs)
            @($groups | Where-Object Root) | Should-BeCollection -Count 0
            @($groups.Jobs) | Should-BeCollection -Count $paths.Count
        }
    }
    It 'never merges drives or UNC shares' {
        $jobs = @('C:\Logs\a.txt', 'C:\Logs\b.txt', 'D:\Logs\only.txt', '\\server\one\a.txt', '\\server\one\b.txt', '\\server\two\only.txt') | ForEach-Object { [pscustomobject]@{ Path = $_ } }
        $groups = @(Get-MailPathGroupHC -Job $jobs)
        $groups | Should-BeCollection -Count 4
        @($groups | Where-Object Root).Root | Should-ContainCollection 'C:\Logs'
        @($groups | Where-Object Root).Root | Should-ContainCollection '\\server\one'
        @($groups | Where-Object { -not $_.Root }) | Should-BeCollection -Count 2
    }
    It 'uses directory boundaries and never treats a selected parent as its own child' {
        $jobs = @('C:\Data', 'C:\Database\a.txt', 'C:\Data\b.txt') | ForEach-Object { [pscustomobject]@{ Path = $_ } }
        foreach ($group in (Get-MailPathGroupHC -Job $jobs)) {
            if ($group.Root) {
                @($group.Jobs.Path | Sort-Object -Unique).Count | Should-BeGreaterThanOrEqual 2
                foreach ($row in $group.Jobs) { $row.Path.StartsWith($group.Root + '\', [StringComparison]::OrdinalIgnoreCase) | Should-BeTrue }
            }
        }
    }
}
Describe 'Outlook path labels' {
    It 'shortens long paths only for Outlook (<Label>)' -ForEach @(
        @{ Label = 'local folder'; Path = 'D:\ConactiveLabShare\CLBEPRD\SAPTMSOrderImport\Backup\Orig_SAPIDOC_005_SALESORDERS'; Name = $null; Tail = 'Orig_SAPIDOC_005_SALESORDERS' }
        @{ Label = 'named UNC file'; Path = '\\server\ConactiveLabShare\Long folder name\Reports & exports\Monthly report.txt'; Name = 'Monthly report'; Tail = 'Monthly report.txt' }
        @{ Label = 'long leaf'; Path = 'D:\Logs\' + ('a' * 80) + '.txt'; Name = $null; Tail = '.txt' }
        @{ Label = 'encoded folder'; Path = 'D:\ConactiveLabShare\Long folder name\Reports & exports\Team ''A''\'; Name = $null; Tail = "Team 'A'" }
    ) {
        $html = Build-MailJobRowHC -Job ([pscustomobject]@{
            Path = $Path; LinkPath = $Path; Name = $Name; Removed = 3; Errors = 1
        })
        $outlook = [regex]::Match($html, '<!--\[if mso\]><span title=''([^'']*)''>(.*?)</span><!\[endif\]-->')
        $label = [System.Net.WebUtility]::HtmlDecode($outlook.Groups[2].Value)
        $label.Length | Should-BeLessThanOrEqual 55
        $label | Should-BeLikeString "...*$Tail"
        [System.Net.WebUtility]::HtmlDecode($outlook.Groups[1].Value) | Should-Be $Path
        $browser = [regex]::Match($html, '<!--\[if !mso\]><!-->(.*?)<!--<!\[endif\]-->').Groups[1].Value
        [System.Net.WebUtility]::HtmlDecode($browser) | Should-Be $Path
        $html.Contains("href='$([System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $Path)))'") | Should-BeTrue
        $html | Should-BeLikeString "*class='removed-count'*>3</td>*class='error-count'*>1</td>*"
    }
    It 'leaves a path of <Length> characters unchanged' -ForEach @(
        @{ Length = 10 }
        @{ Length = 55 }
    ) {
        $path = 'C:\' + ('a' * ($Length - 3))
        $html = Build-MailJobRowHC -Job ([pscustomobject]@{
            Path = $path; LinkPath = $path; Removed = 0; Errors = 0
        })
        $html | Should-NotBeLikeString '*<!--`[if mso`]>*'
        $html | Should-BeLikeString "*>$path</a>*"
    }
}
Describe 'Outlook header count wrapping' {
    It 'keeps count and label together for <Removed> removals and <Errors> errors' -ForEach @(
        @{ Removed = 0; Errors = 0 }
        @{ Removed = 1; Errors = 1 }
        @{ Removed = 2; Errors = 2 }
        @{ Removed = 60013; Errors = 1200 }
    ) {
        $html = Build-MailComputerCardHC -ComputerName 'PC1' -Job @(
            [pscustomobject]@{
                Path = 'C:\Logs'
                LinkPath = '\\PC1\C$\Logs'
                Description = 'Remove all files'
                Removed = $Removed
                Errors = $Errors
            }
        )
        $header = [regex]::Match($html, "<td[^>]*width='140'[^>]*>(.*?)</td>").Groups[1].Value
        $expected = "$Removed&nbsp;removed"
        if ($Errors) {
            $expected += " &middot; $Errors&nbsp;error$(if ($Errors -ne 1) { 's' })"
        }
        $header | Should-Be $expected
    }
}
Describe 'Get-LogFolderHC' {
    It 'creates a folder that does not exist' {
        $testFolder = Join-Path (Get-Item 'TestDrive:\').FullName 'new\sub'

        $actual = Get-LogFolderHC -Path $testFolder

        $actual | Should-Be $testFolder
        (Test-Path -LiteralPath $testFolder) | Should-BeTrue
    }
    It 'returns the path of an existing folder' {
        $testFolder = (New-Item 'TestDrive:\existing' -ItemType Directory).FullName

        Get-LogFolderHC -Path $testFolder | Should-Be $testFolder
    }
    It 'resolves a relative path against the script folder' {
        $testName = "test_$([guid]::NewGuid())"
        $testExpected = Join-Path (Split-Path $PSScriptRoot) $testName

        try {
            Get-LogFolderHC -Path $testName | Should-Be $testExpected
            (Test-Path -LiteralPath $testExpected) | Should-BeTrue
        }
        finally {
            Remove-Item -LiteralPath $testExpected -Force -ErrorAction Ignore
        }
    }
    It 'throws when the folder cannot be created' {
        { Get-LogFolderHC -Path 'x:\notExisting' } |
        Should-Throw "Failed creating log folder 'x:\notExisting'*"
    }
}
Describe 'Get-StringValueHC' {
    It 'returns NULL for an empty value' {
        Get-StringValueHC -Name '' | Should-BeNull
    }
    It 'returns a plain value as is' {
        Get-StringValueHC -Name 'plain' | Should-Be 'plain'
    }
    It 'returns the value of an environment variable' {
        $env:TEST_GET_STRING_VALUE_HC = 'secret'

        try {
            Get-StringValueHC -Name 'ENV:TEST_GET_STRING_VALUE_HC' |
            Should-Be 'secret'
            Get-StringValueHC -Name 'env: TEST_GET_STRING_VALUE_HC' |
            Should-Be 'secret'
        }
        finally {
            Remove-Item -Path 'Env:\TEST_GET_STRING_VALUE_HC'
        }
    }
    It 'throws when the environment variable does not exist' {
        { Get-StringValueHC -Name 'ENV:NOT_EXISTING_VARIABLE_HC' } |
        Should-Throw "Environment variable 'NOT_EXISTING_VARIABLE_HC' not found."
    }
}
Describe 'Invoke-WithOptionalParallelismHC' {
    It 'runs sequentially when ThrottleLimit is <_>' -ForEach @(0, 1) {
        $actual = Invoke-WithOptionalParallelismHC -InputObject 1, 2, 3 -ThrottleLimit $_ -ScriptBlock {
            param($item)
            [PSCustomObject]@{ Value = $item; Runspace = [runspace]::DefaultRunspace.InstanceId }
        }

        $actual.Value | Should-BeCollection @(1, 2, 3)
        $actual.Runspace | Sort-Object -Unique |
        Should-Be ([runspace]::DefaultRunspace.InstanceId)
    }
    It 'runs in parallel runspaces when ThrottleLimit is higher than 1' {
        $actual = Invoke-WithOptionalParallelismHC -InputObject 1, 2, 3 -ThrottleLimit 3 -ScriptBlock {
            param($item)
            [PSCustomObject]@{ Value = $item; Runspace = [runspace]::DefaultRunspace.InstanceId }
        }

        $actual.Value | Sort-Object | Should-BeCollection @(1, 2, 3)
        $actual.Runspace | Should-NotContainCollection ([runspace]::DefaultRunspace.InstanceId)
    }
    It 'passes ArgumentList after the input object when ThrottleLimit is <_>' -ForEach @(1, 2) {
        $actual = Invoke-WithOptionalParallelismHC -InputObject 'a' -ThrottleLimit $_ -ArgumentList 'b', 'c' -ScriptBlock {
            param($item, $second, $third)
            "$item$second$third"
        }

        $actual | Should-Be 'abc'
    }
    It 'returns nothing for an empty input' {
        Invoke-WithOptionalParallelismHC -InputObject @() -ThrottleLimit 2 -ScriptBlock { 'x' } |
        Should-BeCollection -Count 0
    }
}
Describe 'Out-LogFileHC' {
    BeforeEach {
        $testPartialPath = Join-Path (Get-Item 'TestDrive:\').FullName "log_$([guid]::NewGuid())"
        $testData = @(
            [PSCustomObject]@{ DateTime = Get-Date; Message = 'first' }
        )
    }
    It 'creates a .json file and returns its path' {
        $actual = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.json'

        $actual | Should-Be "$testPartialPath.json"
        (Get-Content -LiteralPath $actual -Raw | ConvertFrom-Json).Message |
        Should-Be 'first'
    }
    It 'converts an ErrorRecord message to a string in a .json file' {
        $testError = try { throw 'Oops' } catch { $_ }

        $actual = Out-LogFileHC -DataToExport ([PSCustomObject]@{ DateTime = Get-Date; Message = $testError }) -PartialPath $testPartialPath -FileExtensions '.json'

        (Get-Content -LiteralPath $actual -Raw | ConvertFrom-Json).Message |
        Should-Be 'Oops'
    }
    It 'keeps the existing entries of a .json file with Append' {
        $null = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.json'

        $actual = Out-LogFileHC -DataToExport ([PSCustomObject]@{ DateTime = Get-Date; Message = 'second' }) -PartialPath $testPartialPath -FileExtensions '.json' -Append

        (Get-Content -LiteralPath $actual -Raw | ConvertFrom-Json).Message |
        Should-BeCollection @('second', 'first')
    }
    It 'overwrites a .json file without Append' {
        $null = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.json'

        $actual = Out-LogFileHC -DataToExport ([PSCustomObject]@{ DateTime = Get-Date; Message = 'second' }) -PartialPath $testPartialPath -FileExtensions '.json'

        (Get-Content -LiteralPath $actual -Raw | ConvertFrom-Json).Message |
        Should-Be 'second'
    }
    It 'creates a .txt file' {
        $actual = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.txt'

        $actual | Should-Be "$testPartialPath.txt"
        Get-Content -LiteralPath $actual -Raw | Should-BeLikeString '*Message*first*'
    }
    It 'creates a file for each extension' {
        $actual = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.txt', '.json'

        $actual | Should-BeCollection -Count 2
        $actual | ForEach-Object { (Test-Path -LiteralPath $_) | Should-BeTrue }
    }
    It 'rejects an unsupported extension' {
        { Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.csv' } |
        Should-Throw '*does not belong to the set*'
    }
}
Describe 'Send-MailKitMessageHC' {
    BeforeAll {
        $testParams = @{
            MailKitAssemblyPath = 'x:\MailKit.dll'
            MimeKitAssemblyPath = 'x:\MimeKit.dll'
            SmtpServerName      = 'SMTP_SERVER'
            SmtpPort            = 25
            From                = 'm@example.com'
            To                  = 'bob@example.com'
            Subject             = 'Subject'
            Body                = 'Body'
        }
    }
    It 'throws when To and Bcc are both missing' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Remove('To')

        { Send-MailKitMessageHC @testNewParams } |
        Should-Throw "*Either 'To' to 'Bcc' is required for sending emails*"
    }
    It 'throws when a <Property> address is not valid' -ForEach @(
        @{ Property = 'To' }
        @{ Property = 'Bcc' }
    ) {
        $testNewParams = $testParams.Clone()
        $testNewParams.$Property = 'notAnEmail'

        { Send-MailKitMessageHC @testNewParams } |
        Should-Throw "*$Property email address 'notAnEmail' not valid.*"
    }
    It 'rejects an invalid From address' {
        $testNewParams = $testParams.Clone()
        $testNewParams.From = 'notAnEmail'

        { Send-MailKitMessageHC @testNewParams } |
        Should-Throw "*Cannot validate argument on parameter 'From'*"
    }
    It 'rejects an unsupported SMTP port' {
        $testNewParams = $testParams.Clone()
        $testNewParams.SmtpPort = 26

        { Send-MailKitMessageHC @testNewParams } |
        Should-Throw "*Cannot validate argument on parameter 'SmtpPort'*"
    }
    It 'throws when an assembly cannot be loaded' -Skip:(
        [bool]([AppDomain]::CurrentDomain.GetAssemblies().FullName -like 'MimeKit, *')
    ) {
        { Send-MailKitMessageHC @testParams } |
        Should-Throw "*Failed to load MimeKit assembly 'x:\MimeKit.dll'*"
    }
}
Describe 'Write-EventsToEventLogHC' {
    BeforeAll {
        Mock New-EventLog
        Mock Write-EventLog
    }
    It 'creates the event log source when it does not exist' {
        $testSource = "Test source $([guid]::NewGuid())"

        Write-EventsToEventLogHC -Source $testSource -LogName 'Scripts' -Events @()

        Should-Invoke New-EventLog -Times 1 -Exactly -Scope It -ParameterFilter {
            ($Source -eq $testSource) -and ($LogName -eq 'Scripts')
        }
    }
    It 'writes one event per object with a message built from its properties' {
        $testEvents = @(
            [PSCustomObject]@{ Message = 'one'; EntryType = 'Error'; EventID = '2' }
            [PSCustomObject]@{ Message = 'two'; FileName = 'a.txt' }
        )

        Write-EventsToEventLogHC -Source "Test source $([guid]::NewGuid())" -LogName 'Scripts' -Events $testEvents

        Should-Invoke Write-EventLog -Times 2 -Exactly -Scope It
        Should-Invoke Write-EventLog -Times 1 -Exactly -Scope It -ParameterFilter {
            ($EntryType -eq 'Error') -and ($EventId -eq 2) -and
            ($Message -eq "`n- Message 'one'")
        }
    }
    It 'defaults EntryType to Information and EventID to 4' {
        Write-EventsToEventLogHC -Source "Test source $([guid]::NewGuid())" -LogName 'Scripts' -Events ([PSCustomObject]@{ Message = 'two'; FileName = 'a.txt' })

        Should-Invoke Write-EventLog -Times 1 -Exactly -Scope It -ParameterFilter {
            ($EntryType -eq 'Information') -and ($EventId -eq 4) -and
            ($Message -eq "`n- Message 'two'`n- FileName 'a.txt'")
        }
    }
    It 'throws when the source is registered with another event log' {
        { Write-EventsToEventLogHC -Source 'Application Error' -LogName 'Scripts' -Events @() } |
        Should-Throw "*already registered with event log name 'Application'*"

        Should-NotInvoke New-EventLog -Scope It
    }
}
