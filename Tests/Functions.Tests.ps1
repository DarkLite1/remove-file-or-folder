#Requires -Version 7
#Requires -Modules Pester

BeforeAll {
    . (Join-Path -Path (Split-Path $PSScriptRoot) -ChildPath 'Functions.ps1')
}
Describe 'ConvertTo-FileUrlHC' {
    It 'converts <Path>' -ForEach @(
        @{ Path = '\\server\share\my folder'; Expected = 'file:////server/share/my%20folder' }
        @{ Path = 'C:\Temp'; Expected = 'file://C:/Temp' }
        @{ Path = ''; Expected = '' }
    ) {
        ConvertTo-FileUrlHC -Path $Path | Should -Be $Expected
    }
}
Describe 'New-PillHtmlHC' {
    It 'renders a browser span and an Outlook VML shape' {
        $actual = New-PillHtmlHC -Text 'Error' -Bg '#dc2626'

        $actual | Should -BeLike '*<!--`[if mso`]>*<v:roundrect*fillcolor="#dc2626"*>ERROR</center>*<!`[endif`]-->*'
        $actual | Should -BeLike '*<!--`[if !mso`]><!-->*background-color:#dc2626*>Error</span>*'
    }
    It 'renders nothing without text' {
        New-PillHtmlHC -Text '' -Bg '#dc2626' | Should -BeNullOrEmpty
    }
}
Describe 'Get-TaskDescriptionHC' {
    It 'shows active attribute exclusions' {
        Get-TaskDescriptionHC -Task ([pscustomobject]@{
            Type = 'RemoveFilesInFolder'
            ExcludeAttributes = @('Hidden', 'System')
        }) | Should -Be 'Remove all files, excluding hidden/system items'
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
        Get-TaskDescriptionHC -Task ([PSCustomObject]$Task) | Should -Be $Expected
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
        $actual | Should -BeLike '*<h1>Remove &lt;test&gt;</h1>*'
        $actual | Should -BeLike '*<p>Custom body</p>*'
    }
    It 'has a fixed width for Outlook' {
        $actual | Should -BeLike "*<!--``[if mso``]>*width=`"$($testTheme.BodyWidth)`"*"
    }
    It 'shows the totals as pills' {
        $actual | Should -BeLike '*>4 Removed</span>*'
        $actual | Should -BeLike '*>3 Errors</span>*'
    }
    It 'shows the encoded system errors' {
        $actual | Should -BeLike '*System Errors (1)*Oops &amp; co*'
    }
    It 'links the log folder and mentions the attachments' {
        $actual | Should -BeLike "*href='C:\Log folder'*Open log folder*details in the attachments*"
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
        $outlookLink.Contains("href='$([System.Net.WebUtility]::HtmlEncode($Path))'") | Should -BeTrue
        $outlookLink | Should -Not -BeLike '*target=*'
        $outlookLink | Should -Not -BeLike '*file://*'
        $browserLink.Contains("href='$Url'") | Should -BeTrue
    }
    It 'shows an encoded browser-view link only inside an Outlook conditional row' {
        $testNewParams = $testParams.Clone()
        $testNewParams.BrowserViewFilePath = '\\server\Logs\Team A & B - Mail.html'

        $html = Get-MailBodyHtmlHC @testNewParams

        $html | Should -Match '(?s)<!--\[if mso\]>\s*<tr><td[^>]*><p[^>]*>If this mail is not visible, please <a[^>]+>click here to view it in the browser</a>\.</p></td></tr>\s*<!\[endif\]-->'
        $html | Should -BeLike "*href='file:////server/Logs/Team%20A%20&amp;%20B%20-%20Mail.html'*"
        $html | Should -BeLike '*title="\\server\Logs\Team A &amp; B - Mail.html"*'
    }
    It 'omits the browser-view link without a saved HTML path' {
        $actual | Should -Not -BeLike '*If this mail is not visible*'
    }
    It 'shows a card per computer, the one with errors first and the idle one last' {
        $errIndex = $actual.IndexOf('>PC-ERR</p>')
        $okIndex = $actual.IndexOf('>PC-OK</p>')
        $idleIndex = $actual.IndexOf('>PC-IDLE</p>')

        $errIndex | Should -BeGreaterThan 0
        $errIndex | Should -BeLessThan $okIndex
        $okIndex | Should -BeLessThan $idleIndex
    }
    It 'colors the card headers by result' {
        $actual | Should -BeLike "*linear-gradient(135deg, $($testTheme.GradError[0])*>PC-ERR</p>*"
        $actual | Should -BeLike "*linear-gradient(135deg, $($testTheme.GradSuccess[0])*>PC-OK</p>*"
        $actual | Should -BeLike "*linear-gradient(135deg, $($testTheme.GradIdle[0])*>PC-IDLE</p>*"
    }
    It 'shows the totals in the card header' {
        $actual | Should -BeLike '*>PC-ERR</p>*1&nbsp;removed &middot; 2&nbsp;errors*'
        $actual | Should -BeLike '*>PC-OK</p>*2 paths*3&nbsp;removed*'
    }
    It 'shows the task description above path rows with separate counters' {
        $actual | Should -BeLike '*Remove all files</caption>*'
        $actual | Should -BeLike "*href='file:////PC-OK/D$/Logs'*>Logs</a>*D:\Logs*class='removed-count'*>3</td>*"
        $actual | Should -BeLike "*E:\&lt;Data&gt;*class='removed-count'*>1</td>*class='error-count'*>2</td>*"
    }
    It 'highlights error counters without individual row cards or badges' {
        $actual | Should -BeLike "*class='error-count'*color:$($testTheme.AccentError)*>2</td>*"
        $pathRows = [regex]::Matches($actual, "(?s)<tr class='path-row'>.*?</tr>").Value -join ''
        $pathRows | Should -Not -BeLike '*border-left:3px solid*'
        $actual | Should -Not -BeLike '*>Error</span>*'
    }
    It 'renders the description once per task and keeps each path counter' {
        $html = Build-MailComputerCardHC -ComputerName 'PC1' -Job @(
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\First'; LinkPath = 'C:\First'; Description = 'Remove files & folders'; Removed = 5; Errors = 0 }
            [pscustomobject]@{ TaskIndex = 0; Path = 'C:\Second'; LinkPath = 'C:\Second'; Description = 'Remove files & folders'; Removed = 0; Errors = 2 }
            [pscustomobject]@{ TaskIndex = 1; Path = 'C:\First'; LinkPath = 'C:\First'; Description = 'Remove files & folders'; Removed = 0; Errors = 0 }
        )
        [regex]::Matches($html, 'Remove files &amp; folders').Count | Should -Be 2
        [regex]::Matches($html, 'class="task-table"').Count | Should -Be 2
        $pathRows = [regex]::Matches($html, "(?s)<tr class='path-row'>.*?</tr>").Value
        $pathRows | Should -HaveCount 3
        $pathRows[0] | Should -BeLike "*C:\Second*class='removed-count'*>0</td>*class='error-count'*>2</td>*"
        $pathRows[1] | Should -BeLike "*C:\First*class='removed-count'*>5</td>*class='error-count'*>0</td>*"
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
            $row.td | Should -HaveCount 3
            foreach ($cell in $row.td) {
                if ($errorCount) {
                    $cell.bgcolor | Should -Be $testTheme.StatusError
                    $cell.style | Should -BeLike "*background-color:$($testTheme.StatusError)*color:$($testTheme.AccentError)*"
                }
                else {
                    $cell.bgcolor | Should -Be $testTheme.BgWhite
                    $cell.style | Should -Not -BeLike "*color:$($testTheme.AccentError)*"
                }
            }
            foreach ($label in $row.SelectNodes('.//div | .//a')) {
                if ($errorCount) { $label.style | Should -BeLike "*color:$($testTheme.AccentError)*" }
                else { $label.style | Should -Not -BeLike "*color:$($testTheme.AccentError)*" }
            }
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
        $pathRows | Should -HaveCount 5
        $pathRows[0] | Should -BeLike '*C:\Charlie*'
        $pathRows[1] | Should -BeLike '*C:\Delta*'
        $pathRows[2] | Should -BeLike '*C:\Alpha*'
        $pathRows[3] | Should -BeLike '*C:\Bravo*'
        $pathRows[4] | Should -BeLike '*C:\Zulu*'
    }
    It 'shows the run times in the footer' {
        $actual | Should -BeLike '*Started*02/10/2026 07:00*Ended*02/10/2026 08:30*Duration*01:30:15*'
    }
    It 'shows a message when there are no jobs' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Job = @()

        Get-MailBodyHtmlHC @testNewParams | Should -BeLike '*No tasks were executed.*'
    }
    It 'shows a grey removed pill and no error pill when nothing happened' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Removed = 0
        $testNewParams.Errors = 0

        $result = Get-MailBodyHtmlHC @testNewParams

        $result | Should -BeLike "*background-color:$($testTheme.AccentIdle)*>0 Removed</span>*"
        $result | Should -Not -BeLike '*Errors</span>*'
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
        $label.Length | Should -BeLessOrEqual 55
        $label | Should -BeLike "...*$Tail"
        [System.Net.WebUtility]::HtmlDecode($outlook.Groups[1].Value) | Should -Be $Path
        $browser = [regex]::Match($html, '<!--\[if !mso\]><!-->(.*?)<!--<!\[endif\]-->').Groups[1].Value
        [System.Net.WebUtility]::HtmlDecode($browser) | Should -Be $Path
        $html.Contains("href='$([System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $Path)))'") | Should -BeTrue
        $html | Should -BeLike "*class='removed-count'*>3</td>*class='error-count'*>1</td>*"
    }
    It 'leaves a path of <Length> characters unchanged' -ForEach @(
        @{ Length = 10 }
        @{ Length = 55 }
    ) {
        $path = 'C:\' + ('a' * ($Length - 3))
        $html = Build-MailJobRowHC -Job ([pscustomobject]@{
            Path = $path; LinkPath = $path; Removed = 0; Errors = 0
        })
        $html | Should -Not -BeLike '*<!--`[if mso`]>*'
        $html | Should -BeLike "*>$path</a>*"
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
        $header | Should -Be $expected
    }
}
Describe 'Get-LogFolderHC' {
    It 'creates a folder that does not exist' {
        $testFolder = Join-Path (Get-Item 'TestDrive:\').FullName 'new\sub'

        $actual = Get-LogFolderHC -Path $testFolder

        $actual | Should -Be $testFolder
        $testFolder | Should -Exist
    }
    It 'returns the path of an existing folder' {
        $testFolder = (New-Item 'TestDrive:\existing' -ItemType Directory).FullName

        Get-LogFolderHC -Path $testFolder | Should -Be $testFolder
    }
    It 'resolves a relative path against the script folder' {
        $testName = "test_$([guid]::NewGuid())"
        $testExpected = Join-Path (Split-Path $PSScriptRoot) $testName

        try {
            Get-LogFolderHC -Path $testName | Should -Be $testExpected
            $testExpected | Should -Exist
        }
        finally {
            Remove-Item -LiteralPath $testExpected -Force -ErrorAction Ignore
        }
    }
    It 'throws when the folder cannot be created' {
        { Get-LogFolderHC -Path 'x:\notExisting' } |
        Should -Throw "Failed creating log folder 'x:\notExisting'*"
    }
}
Describe 'Get-StringValueHC' {
    It 'returns NULL for an empty value' {
        Get-StringValueHC -Name '' | Should -BeNullOrEmpty
    }
    It 'returns a plain value as is' {
        Get-StringValueHC -Name 'plain' | Should -Be 'plain'
    }
    It 'returns the value of an environment variable' {
        $env:TEST_GET_STRING_VALUE_HC = 'secret'

        try {
            Get-StringValueHC -Name 'ENV:TEST_GET_STRING_VALUE_HC' |
            Should -Be 'secret'
            Get-StringValueHC -Name 'env: TEST_GET_STRING_VALUE_HC' |
            Should -Be 'secret'
        }
        finally {
            Remove-Item -Path 'Env:\TEST_GET_STRING_VALUE_HC'
        }
    }
    It 'throws when the environment variable does not exist' {
        { Get-StringValueHC -Name 'ENV:NOT_EXISTING_VARIABLE_HC' } |
        Should -Throw "Environment variable 'NOT_EXISTING_VARIABLE_HC' not found."
    }
}
Describe 'Invoke-WithOptionalParallelismHC' {
    It 'runs sequentially when ThrottleLimit is <_>' -ForEach @(0, 1) {
        $actual = Invoke-WithOptionalParallelismHC -InputObject 1, 2, 3 -ThrottleLimit $_ -ScriptBlock {
            param($item)
            [PSCustomObject]@{ Value = $item; Runspace = [runspace]::DefaultRunspace.InstanceId }
        }

        $actual.Value | Should -Be @(1, 2, 3)
        $actual.Runspace | Sort-Object -Unique |
        Should -Be ([runspace]::DefaultRunspace.InstanceId)
    }
    It 'runs in parallel runspaces when ThrottleLimit is higher than 1' {
        $actual = Invoke-WithOptionalParallelismHC -InputObject 1, 2, 3 -ThrottleLimit 3 -ScriptBlock {
            param($item)
            [PSCustomObject]@{ Value = $item; Runspace = [runspace]::DefaultRunspace.InstanceId }
        }

        $actual.Value | Sort-Object | Should -Be @(1, 2, 3)
        $actual.Runspace | Should -Not -Contain ([runspace]::DefaultRunspace.InstanceId)
    }
    It 'passes ArgumentList after the input object when ThrottleLimit is <_>' -ForEach @(1, 2) {
        $actual = Invoke-WithOptionalParallelismHC -InputObject 'a' -ThrottleLimit $_ -ArgumentList 'b', 'c' -ScriptBlock {
            param($item, $second, $third)
            "$item$second$third"
        }

        $actual | Should -Be 'abc'
    }
    It 'returns nothing for an empty input' {
        Invoke-WithOptionalParallelismHC -InputObject @() -ThrottleLimit 2 -ScriptBlock { 'x' } |
        Should -BeNullOrEmpty
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

        $actual | Should -Be "$testPartialPath.json"
        (Get-Content -LiteralPath $actual -Raw | ConvertFrom-Json).Message |
        Should -Be 'first'
    }
    It 'converts an ErrorRecord message to a string in a .json file' {
        $testError = try { throw 'Oops' } catch { $_ }

        $actual = Out-LogFileHC -DataToExport ([PSCustomObject]@{ DateTime = Get-Date; Message = $testError }) -PartialPath $testPartialPath -FileExtensions '.json'

        (Get-Content -LiteralPath $actual -Raw | ConvertFrom-Json).Message |
        Should -Be 'Oops'
    }
    It 'keeps the existing entries of a .json file with Append' {
        $null = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.json'

        $actual = Out-LogFileHC -DataToExport ([PSCustomObject]@{ DateTime = Get-Date; Message = 'second' }) -PartialPath $testPartialPath -FileExtensions '.json' -Append

        (Get-Content -LiteralPath $actual -Raw | ConvertFrom-Json).Message |
        Should -Be @('second', 'first')
    }
    It 'overwrites a .json file without Append' {
        $null = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.json'

        $actual = Out-LogFileHC -DataToExport ([PSCustomObject]@{ DateTime = Get-Date; Message = 'second' }) -PartialPath $testPartialPath -FileExtensions '.json'

        (Get-Content -LiteralPath $actual -Raw | ConvertFrom-Json).Message |
        Should -Be 'second'
    }
    It 'creates a .txt file' {
        $actual = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.txt'

        $actual | Should -Be "$testPartialPath.txt"
        Get-Content -LiteralPath $actual -Raw | Should -BeLike '*Message*first*'
    }
    It 'creates a file for each extension' {
        $actual = Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.txt', '.json'

        $actual | Should -HaveCount 2
        $actual | ForEach-Object { $_ | Should -Exist }
    }
    It 'rejects an unsupported extension' {
        { Out-LogFileHC -DataToExport $testData -PartialPath $testPartialPath -FileExtensions '.csv' } |
        Should -Throw '*does not belong to the set*'
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
        Should -Throw "*Either 'To' to 'Bcc' is required for sending emails*"
    }
    It 'throws when a <Property> address is not valid' -ForEach @(
        @{ Property = 'To' }
        @{ Property = 'Bcc' }
    ) {
        $testNewParams = $testParams.Clone()
        $testNewParams.$Property = 'notAnEmail'

        { Send-MailKitMessageHC @testNewParams } |
        Should -Throw "*$Property email address 'notAnEmail' not valid.*"
    }
    It 'rejects an invalid From address' {
        $testNewParams = $testParams.Clone()
        $testNewParams.From = 'notAnEmail'

        { Send-MailKitMessageHC @testNewParams } |
        Should -Throw "*Cannot validate argument on parameter 'From'*"
    }
    It 'rejects an unsupported SMTP port' {
        $testNewParams = $testParams.Clone()
        $testNewParams.SmtpPort = 26

        { Send-MailKitMessageHC @testNewParams } |
        Should -Throw "*Cannot validate argument on parameter 'SmtpPort'*"
    }
    It 'throws when an assembly cannot be loaded' -Skip:(
        [bool]([AppDomain]::CurrentDomain.GetAssemblies().FullName -like 'MimeKit, *')
    ) {
        { Send-MailKitMessageHC @testParams } |
        Should -Throw "*Failed to load MimeKit assembly 'x:\MimeKit.dll'*"
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

        Should -Invoke New-EventLog -Times 1 -Exactly -Scope It -ParameterFilter {
            ($Source -eq $testSource) -and ($LogName -eq 'Scripts')
        }
    }
    It 'writes one event per object with a message built from its properties' {
        $testEvents = @(
            [PSCustomObject]@{ Message = 'one'; EntryType = 'Error'; EventID = '2' }
            [PSCustomObject]@{ Message = 'two'; FileName = 'a.txt' }
        )

        Write-EventsToEventLogHC -Source "Test source $([guid]::NewGuid())" -LogName 'Scripts' -Events $testEvents

        Should -Invoke Write-EventLog -Times 2 -Exactly -Scope It
        Should -Invoke Write-EventLog -Times 1 -Exactly -Scope It -ParameterFilter {
            ($EntryType -eq 'Error') -and ($EventId -eq 2) -and
            ($Message -eq "`n- Message 'one'")
        }
    }
    It 'defaults EntryType to Information and EventID to 4' {
        Write-EventsToEventLogHC -Source "Test source $([guid]::NewGuid())" -LogName 'Scripts' -Events ([PSCustomObject]@{ Message = 'two'; FileName = 'a.txt' })

        Should -Invoke Write-EventLog -Times 1 -Exactly -Scope It -ParameterFilter {
            ($EntryType -eq 'Information') -and ($EventId -eq 4) -and
            ($Message -eq "`n- Message 'two'`n- FileName 'a.txt'")
        }
    }
    It 'throws when the source is registered with another event log' {
        { Write-EventsToEventLogHC -Source 'Application Error' -LogName 'Scripts' -Events @() } |
        Should -Throw "*already registered with event log name 'Application'*"

        Should -Not -Invoke New-EventLog -Scope It
    }
}
