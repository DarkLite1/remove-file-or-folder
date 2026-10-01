#Requires -Version 7
#Requires -Modules Pester

BeforeAll {
    . (Join-Path -Path (Split-Path $PSScriptRoot) -ChildPath 'Functions.ps1')
}
Describe 'ConvertTo-HtmlListHC' {
    It 'creates a list item for every message' {
        $actual = 'a', 'b' | ConvertTo-HtmlListHC

        $actual | Should -BeLike '*<ul>*<li style="margin: 10px 0;">a</li>*<li style="margin: 10px 0;">b</li>*</ul>*'
    }
    It 'adds a header and a foot note when given' {
        $actual = ConvertTo-HtmlListHC -Message 'a' -Header 'Title' -FootNote 'Note'

        $actual | Should -BeLike '*<h3>Title</h3>*'
        $actual | Should -BeLike '*<i><font size="2">* Note</font></i>*'
    }
    It 'adds no header or foot note when not given' {
        $actual = ConvertTo-HtmlListHC -Message 'a'

        $actual | Should -Not -BeLike '*<h3>*'
        $actual | Should -Not -BeLike '*<font*'
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
