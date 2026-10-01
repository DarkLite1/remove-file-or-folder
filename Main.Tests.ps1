#Requires -Version 7
#Requires -Modules Pester
#Requires -Modules ImportExcel

BeforeAll {
    $testLogFolder = (New-Item 'TestDrive:/log' -ItemType Directory).FullName

    $testInputFile = @{
        MaxConcurrent = @{
            JobsTotal       = 1
            JobsPerComputer = 1
        }
        Remove        = @{
            File          = @(
                @{
                    Name         = 'FTP log file'
                    ComputerName = 'PC1'
                    Path         = 'z:\file.txt'
                    OlderThan    = @{
                        Quantity = 1
                        Unit     = 'Day'
                    }
                }
            )
            FilesInFolder = @(
                @{
                    Name         = 'App log folder'
                    ComputerName = 'PC2'
                    Path         = 'z:\folder'
                    Recurse      = $true
                    OlderThan    = @{
                        Quantity = 1
                        Unit     = 'Day'
                    }
                }
            )
            EmptyFolders  = @(
                @{
                    Name         = 'Delivery notes'
                    ComputerName = 'PC3'
                    Path         = 'z:\folder'
                }
            )
        }
        Settings      = @{
            ScriptName     = 'Test (Brecht)'
            SendMail       = @{
                When         = 'Always'
                From         = 'm@example.com'
                To           = 'bob@contoso.com'
                Bcc          = 'admin@contoso.com'
                Subject      = $null
                Body         = 'Email body'
                Smtp         = @{
                    ServerName     = 'SMTP_SERVER'
                    Port           = 25
                    ConnectionType = 'StartTls'
                    UserName       = 'bob'
                    Password       = 'pass'
                }
                AssemblyPath = @{
                    MailKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.13.0\lib\net8.0\MailKit.dll'
                    MimeKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.13.0\lib\net8.0\MimeKit.dll'
                }
            }
            SaveLogFiles   = @{
                Where               = @{
                    Folder = $testLogFolder
                }
                DeleteLogsAfterDays = 1
            }
            SaveInEventLog = @{
                Save    = $true
                LogName = 'Scripts'
            }
        }
    }

    $testOutParams = @{
        FilePath = (New-Item 'TestDrive:/Test.json' -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testData = @(
        @{
            DateTime     = Get-Date
            ComputerName = $testInputFile.Remove.File[0].ComputerName
            Type         = 'File'
            FullName     = 'z:\file1.txt'
            CreationTime = Get-Date
            Action       = 'Removed'
            Error        = $null
        }
        @{
            DateTime     = Get-Date
            ComputerName = $testInputFile.Remove.FilesInFolder[0].ComputerName
            Type         = 'File'
            FullName     = 'z:\file2.txt'
            CreationTime = Get-Date
            Action       = $null
            Error        = 'File in use'
        }
        @{
            DateTime     = Get-Date
            ComputerName = $testInputFile.Remove.FilesInFolder[0].ComputerName
            Type         = 'File'
            FullName     = 'z:\file3.txt'
            CreationTime = Get-Date
            Action       = 'Removed'
            Error        = $null
        }
        @{
            DateTime     = Get-Date
            ComputerName = $testInputFile.Remove.EmptyFolders[0].ComputerName
            Type         = 'EmptyFolder'
            FullName     = 'z:\folder'
            CreationTime = Get-Date
            Action       = 'Removed'
            Error        = $null
        }
    )

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ConfigurationJsonFile = $testOutParams.FilePath
        Path                  = @{
            RemoveFileScript          = (New-Item 'TestDrive:/b.ps1' -ItemType File).FullName
            RemoveEmptyFoldersScript  = (New-Item 'TestDrive:/a.ps1' -ItemType File).FullName
            RemoveFilesInFolderScript = (New-Item 'TestDrive:/c.ps1' -ItemType File).FullName
        }
    }

    function Copy-ObjectHC {
        param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $InputObject | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    }

    function Test-NewJsonFileHC {
        param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $InputObject | ConvertTo-Json -Depth 7 | Out-File @testOutParams
    }

    function Clear-TestLogFolderHC {
        Get-ChildItem -LiteralPath $testLogFolder | Remove-Item -Recurse -Force
    }

    function Get-TestSystemErrorsHC {
        $testLogFile = Get-ChildItem -LiteralPath $testLogFolder -File -Filter '* - System errors log.json'

        if (@($testLogFile).Count -ne 1) {
            throw "Expected 1 system errors log file in '$testLogFolder', found $(@($testLogFile).Count)"
        }

        Get-Content -LiteralPath $testLogFile.FullName -Raw | ConvertFrom-Json
    }

    function Get-TestExcelFileHC {
        Get-ChildItem -LiteralPath $testLogFolder -File -Filter '*.xlsx'
    }

    function Send-MailKitMessageHC {
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
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
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
    }

    Mock Invoke-Command
    Mock New-PSSession {
        New-MockObject -Type 'System.Management.Automation.Runspaces.PSSession'
    }
    Mock Remove-PSSession
    Mock Start-Sleep
    Mock Send-MailKitMessageHC
    Mock New-EventLog
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ConfigurationJsonFile') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'an incorrect input file' {
    BeforeEach {
        Clear-TestLogFolderHC
    }
    It 'ConfigurationJsonFile not found' {
        $testNewParams = $testParams.Clone()
        $testNewParams.ConfigurationJsonFile = 'nonExisting.json'

        .$testScript @testNewParams -WarningVariable testWarnings -WarningAction SilentlyContinue

        $LASTEXITCODE | Should -Be 1
        ($testWarnings -join "`n") | Should -BeLike '*nonExisting.json*'
        Should -Not -Invoke Invoke-Command -Scope It
    }
    It 'Path.<_> not found' -ForEach @(
        'RemoveEmptyFoldersScript', 'RemoveFileScript', 'RemoveFilesInFolderScript'
    ) {
        Test-NewJsonFileHC $testInputFile

        $testNewParams = $testParams.Clone()
        $testNewParams.Path = $testParams.Path.Clone()
        $testNewParams.Path.$_ = 'c:\NotExisting.ps1'

        .$testScript @testNewParams

        $LASTEXITCODE | Should -Be 1
        ((Get-TestSystemErrorsHC).Message -join "`n") |
        Should -BeLike "*Path.$_ 'c:\NotExisting.ps1' not found*"
        Should -Not -Invoke Invoke-Command -Scope It
    }
    It '<Description>' -ForEach @(
        @{
            Description = 'MaxConcurrent missing'
            Change      = { param($f) $f.PSObject.Properties.Remove('MaxConcurrent') }
            Message     = "Property 'MaxConcurrent' not found"
        }
        @{
            Description = 'Remove missing'
            Change      = { param($f) $f.PSObject.Properties.Remove('Remove') }
            Message     = "Property 'Remove' not found"
        }
        @{
            Description = 'MaxConcurrent.JobsTotal missing'
            Change      = { param($f) $f.MaxConcurrent.JobsTotal = $null }
            Message     = "Property 'MaxConcurrent.JobsTotal' not found"
        }
        @{
            Description = 'MaxConcurrent.JobsPerComputer missing'
            Change      = { param($f) $f.MaxConcurrent.JobsPerComputer = $null }
            Message     = "Property 'MaxConcurrent.JobsPerComputer' not found"
        }
        @{
            Description = 'MaxConcurrent.JobsTotal not a number'
            Change      = { param($f) $f.MaxConcurrent.JobsTotal = 'a' }
            Message     = "Property 'MaxConcurrent.JobsTotal' needs to be a number of 1 or higher, the value 'a' is not supported."
        }
        @{
            Description = 'MaxConcurrent.JobsPerComputer is 0'
            Change      = { param($f) $f.MaxConcurrent.JobsPerComputer = 0 }
            Message     = "Property 'MaxConcurrent.JobsPerComputer' needs to be a number of 1 or higher, the value '0' is not supported."
        }
        @{
            Description = 'Settings.ScriptName missing'
            Change      = { param($f) $f.Settings.ScriptName = $null }
            Message     = "Property 'Settings.ScriptName' not found"
        }
        @{
            Description = 'Settings.SendMail.When missing'
            Change      = { param($f) $f.Settings.SendMail.When = $null }
            Message     = "Property 'Settings.SendMail.When' not found"
        }
        @{
            Description = 'Settings.SendMail.When not supported'
            Change      = { param($f) $f.Settings.SendMail.When = 'Sometimes' }
            Message     = "Property 'Settings.SendMail.When' with value 'Sometimes' is not supported*"
        }
        @{
            Description = 'Settings.SendMail.From missing'
            Change      = { param($f) $f.Settings.SendMail.From = $null }
            Message     = "Property 'Settings.SendMail.From' not found"
        }
        @{
            Description = 'Settings.SendMail.Smtp.ServerName missing'
            Change      = { param($f) $f.Settings.SendMail.Smtp.ServerName = $null }
            Message     = "Property 'Settings.SendMail.Smtp.ServerName' not found"
        }
        @{
            Description = 'Settings.SendMail.To and Bcc missing'
            Change      = { param($f) $f.Settings.SendMail.To = $null; $f.Settings.SendMail.Bcc = $null }
            Message     = "Property 'Settings.SendMail.To' or 'Settings.SendMail.Bcc' not found"
        }
        @{
            Description = 'Settings.SendMail.Smtp.Port not supported'
            Change      = { param($f) $f.Settings.SendMail.Smtp.Port = 26 }
            Message     = "Property 'Settings.SendMail.Smtp.Port' with value '26' is not supported*"
        }
        @{
            Description = 'Settings.SendMail.Smtp.ConnectionType not supported'
            Change      = { param($f) $f.Settings.SendMail.Smtp.ConnectionType = 'Wrong' }
            Message     = "Property 'Settings.SendMail.Smtp.ConnectionType' with value 'Wrong' is not supported*"
        }
        @{
            Description = 'Settings.SaveLogFiles.DeleteLogsAfterDays not a number'
            Change      = { param($f) $f.Settings.SaveLogFiles.DeleteLogsAfterDays = 'abc' }
            Message     = "Property 'Settings.SaveLogFiles.DeleteLogsAfterDays' needs to be a positive number, the value 'abc' is not supported."
        }
        @{
            Description = 'Settings.SaveInEventLog.Save missing'
            Change      = { param($f) $f.Settings.SaveInEventLog.Save = $null }
            Message     = "Property 'Settings.SaveInEventLog.Save' not found"
        }
        @{
            Description = 'Settings.SaveInEventLog.Save not a boolean'
            Change      = { param($f) $f.Settings.SaveInEventLog.Save = 'yes' }
            Message     = "Property 'Settings.SaveInEventLog.Save' needs to be true or false, the value 'yes' is not supported."
        }
        @{
            Description = 'Settings.SaveInEventLog.LogName missing when Save is true'
            Change      = { param($f) $f.Settings.SaveInEventLog.LogName = $null }
            Message     = "Property 'Settings.SaveInEventLog.LogName' not found"
        }
        @{
            Description = 'Remove.File.Path missing'
            Change      = { param($f) $f.Remove.File[0].Path = $null }
            Message     = "Property 'Remove.File.Path' not found"
        }
        @{
            Description = 'Remove.File.OlderThan missing'
            Change      = { param($f) $f.Remove.File[0].OlderThan = $null }
            Message     = "Property 'Remove.File.OlderThan' not found"
        }
        @{
            Description = 'Remove.File.OlderThan.Unit missing'
            Change      = { param($f) $f.Remove.File[0].OlderThan.PSObject.Properties.Remove('Unit') }
            Message     = "No 'Remove.File.OlderThan.Unit' found"
        }
        @{
            Description = 'Remove.File.OlderThan.Unit not supported'
            Change      = { param($f) $f.Remove.File[0].OlderThan.Unit = 'notSupported' }
            Message     = "Value 'notSupported' is not supported by 'Remove.File.OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'."
        }
        @{
            Description = 'Remove.File.OlderThan.Quantity missing'
            Change      = { param($f) $f.Remove.File[0].OlderThan.PSObject.Properties.Remove('Quantity') }
            Message     = "Property 'Remove.File.OlderThan.Quantity' not found. Use value number '0' to move all files."
        }
        @{
            Description = 'Remove.File.OlderThan.Quantity not a number'
            Change      = { param($f) $f.Remove.File[0].OlderThan.Quantity = 'a' }
            Message     = "Property 'Remove.File.OlderThan.Quantity' needs to be a number, the value 'a' is not supported*"
        }
        @{
            Description = 'Remove.File local path without ComputerName'
            Change      = { param($f) $f.Remove.File[0].ComputerName = $null; $f.Remove.File[0].Path = 'd:\bla' }
            Message     = "No 'Remove.File.ComputerName' found for path 'd:\bla'"
        }
        @{
            Description = 'Remove.FilesInFolder.Path missing'
            Change      = { param($f) $f.Remove.FilesInFolder[0].Path = $null }
            Message     = "Property 'Remove.FilesInFolder.Path' not found"
        }
        @{
            Description = 'Remove.FilesInFolder.OlderThan.Unit not supported'
            Change      = { param($f) $f.Remove.FilesInFolder[0].OlderThan.Unit = 'notSupported' }
            Message     = "Value 'notSupported' is not supported by 'Remove.FilesInFolder.OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'."
        }
        @{
            Description = 'Remove.FilesInFolder.OlderThan.Quantity not a number'
            Change      = { param($f) $f.Remove.FilesInFolder[0].OlderThan.Quantity = 'a' }
            Message     = "Property 'Remove.FilesInFolder.OlderThan.Quantity' needs to be a number, the value 'a' is not supported*"
        }
        @{
            Description = 'Remove.FilesInFolder.Recurse not a boolean'
            Change      = { param($f) $f.Remove.FilesInFolder[0].Recurse = 'a' }
            Message     = "Property 'Remove.FilesInFolder.Recurse' is not a boolean value"
        }
        @{
            Description = 'Remove.FilesInFolder local path without ComputerName'
            Change      = { param($f) $f.Remove.FilesInFolder[0].ComputerName = $null; $f.Remove.FilesInFolder[0].Path = 'd:\bla' }
            Message     = "No 'Remove.FilesInFolder.ComputerName' found for path 'd:\bla'"
        }
        @{
            Description = 'Remove.EmptyFolders.Path missing'
            Change      = { param($f) $f.Remove.EmptyFolders[0].Path = $null }
            Message     = "Property 'Remove.EmptyFolders.Path' not found"
        }
        @{
            Description = 'Remove.EmptyFolders local path without ComputerName'
            Change      = { param($f) $f.Remove.EmptyFolders[0].ComputerName = $null; $f.Remove.EmptyFolders[0].Path = 'd:\bla' }
            Message     = "No 'Remove.EmptyFolders.ComputerName' found for path 'd:\bla'"
        }
        @{
            Description = 'there is nothing to execute'
            Change      = { param($f) $f.Remove = [PSCustomObject]@{ File = @() } }
            Message     = 'No tasks to execute'
        }
    ) {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        & $Change $testNewInputFile

        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testParams -WarningVariable testWarnings -WarningAction SilentlyContinue

        $LASTEXITCODE | Should -Be 1
        ($testWarnings -join "`n") | Should -BeLike "*$Message*"
        ((Get-TestSystemErrorsHC).Message -join "`n") | Should -BeLike "*$Message*"
        Should -Not -Invoke Invoke-Command -Scope It
    }
    It 'is reported by e-mail and in the event log' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Remove.File[0].Path = $null

        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testParams

        Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It -ParameterFilter {
            ($Priority -eq 'High') -and
            ($Body -like "*Property 'Remove.File.Path' not found*")
        }
        Should -Invoke Write-EventLog -Scope It -ParameterFilter {
            ($EntryType -eq 'Error') -and
            ($Message -like "*Property 'Remove.File.Path' not found*")
        }
    }
    It 'Settings.SendMail properties are not needed when SendMail.When is Never' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SendMail = [PSCustomObject]@{ When = 'Never' }

        Test-NewJsonFileHC $testNewInputFile

        $global:LASTEXITCODE = 0

        .$testScript @testParams

        $LASTEXITCODE | Should -Be 0
        Should -Invoke Invoke-Command -Times 3 -Exactly -Scope It
        Should -Not -Invoke Send-MailKitMessageHC -Scope It
    }
}
Describe 'execute script' {
    Context 'Remove.File' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Remove = [PSCustomObject]@{
                File = $testNewInputFile.Remove.File
            }

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should -Invoke New-PSSession -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq $testNewInputFile.Remove.File[0].ComputerName) -and
                ($ConfigurationName -eq 'PowerShell.7')
            }
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Session) -and
                ($FilePath -eq $testParams.Path.RemoveFileScript) -and
                ($ArgumentList[0] -eq $testNewInputFile.Remove.File[0].Path) -and
                ($ArgumentList[1] -eq $testNewInputFile.Remove.File[0].OlderThan.Unit) -and
                ($ArgumentList[2] -eq $testNewInputFile.Remove.File[0].OlderThan.Quantity)
            }
        }
    }
    Context 'Remove.FilesInFolder' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Remove = [PSCustomObject]@{
                FilesInFolder = $testNewInputFile.Remove.FilesInFolder
            }

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should -Invoke New-PSSession -Times 1 -Exactly -Scope Context -ParameterFilter {
                $ComputerName -eq $testNewInputFile.Remove.FilesInFolder[0].ComputerName
            }
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Session) -and
                ($FilePath -eq $testParams.Path.RemoveFilesInFolderScript) -and
                ($ArgumentList[0] -eq $testNewInputFile.Remove.FilesInFolder[0].Path) -and
                ($ArgumentList[1] -eq $testNewInputFile.Remove.FilesInFolder[0].OlderThan.Unit) -and
                ($ArgumentList[2] -eq $testNewInputFile.Remove.FilesInFolder[0].OlderThan.Quantity) -and
                ($ArgumentList[3] -eq $testNewInputFile.Remove.FilesInFolder[0].Recurse)
            }
        }
    }
    Context 'Remove.EmptyFolders' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Remove = [PSCustomObject]@{
                EmptyFolders = $testNewInputFile.Remove.EmptyFolders
            }

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should -Invoke New-PSSession -Times 1 -Exactly -Scope Context -ParameterFilter {
                $ComputerName -eq $testNewInputFile.Remove.EmptyFolders[0].ComputerName
            }
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Session) -and
                ($FilePath -eq $testParams.Path.RemoveEmptyFoldersScript) -and
                ($ArgumentList[0] -eq $testNewInputFile.Remove.EmptyFolders[0].Path)
            }
        }
        It 'and close the session' {
            Should -Invoke Remove-PSSession -Times 1 -Exactly -Scope Context
        }
    }
    Context 'PSSessionConfiguration' {
        It 'is used for the remote session' {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Remove = [PSCustomObject]@{
                File = $testNewInputFile.Remove.File
            }
            $testNewInputFile | Add-Member -NotePropertyName 'PSSessionConfiguration' -NotePropertyValue 'PowerShell.7.5'

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should -Invoke New-PSSession -Times 1 -Exactly -Scope It -ParameterFilter {
                $ConfigurationName -eq 'PowerShell.7.5'
            }
        }
    }
}
Describe 'retry a remote job' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Remove = [PSCustomObject]@{
            File = $testNewInputFile.Remove.File
        }

        Test-NewJsonFileHC $testNewInputFile

        $testTransientError = 'Processing data from remote server PC1 failed: The I/O operation has been aborted because of either a thread exit or an application request.'
    }
    BeforeEach {
        Clear-TestLogFolderHC
    }
    It 'up to 3 times on a transient WinRM abort and report the last error' {
        Mock Invoke-Command { throw $testTransientError }

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 3 -Exactly
        Should -Invoke New-PSSession -Times 3 -Exactly
        Should -Invoke Remove-PSSession -Times 3 -Exactly
        Should -Invoke Start-Sleep -Times 2 -Exactly

        $actual = Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName 'Errors'
        $actual.Error | Should -BeLike '*I/O operation has been aborted*'
    }
    It 'and keep the result when a retry succeeds' {
        $script:testAttempt = 0
        Mock Invoke-Command {
            $script:testAttempt++
            if ($script:testAttempt -eq 1) { throw $testTransientError }
            $testData[0]
        }

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 2 -Exactly

        $testExcelFile = Get-TestExcelFileHC
        $actual = Import-Excel -Path $testExcelFile.FullName -WorksheetName 'Overview'
        $actual.Path | Should -Be $testData[0].FullName
        Get-ExcelSheetInfo -Path $testExcelFile.FullName |
        Where-Object Name -EQ 'Errors' | Should -BeNullOrEmpty
    }
    It 'not on other errors' {
        Mock Invoke-Command { throw 'Oops' }

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly
        Should -Invoke Remove-PSSession -Times 1 -Exactly
        Should -Not -Invoke Start-Sleep
    }
}
Describe 'MaxConcurrent' {
    BeforeAll {
        $testJobLogFolder = (New-Item 'TestDrive:/jobLog' -ItemType Directory).FullName

        $testJobScript = (New-Item 'TestDrive:/job.ps1' -ItemType File).FullName
        Set-Content -LiteralPath $testJobScript -Value @"
param(`$Path, `$Unit, `$Quantity)
`$start = [DateTime]::UtcNow.Ticks
Start-Sleep -Milliseconds 1500
`$end = [DateTime]::UtcNow.Ticks
Set-Content -LiteralPath (Join-Path '$testJobLogFolder' ([guid]::NewGuid())) -Value "`$start;`$end"
"@

        $testNewParams = $testParams.Clone()
        $testNewParams.Path = $testParams.Path.Clone()
        $testNewParams.Path.RemoveFileScript = $testJobScript

        function Get-MaxOverlapHC {
            $jobs = Get-ChildItem -LiteralPath $testJobLogFolder -File |
            ForEach-Object {
                $start, $end = (Get-Content -LiteralPath $_.FullName) -split ';'
                [PSCustomObject]@{ Start = [long]$start; End = [long]$end }
            }

            ($jobs | ForEach-Object {
                $job = $_
                @($jobs | Where-Object {
                        ($_.Start -le $job.Start) -and ($_.End -gt $job.Start)
                    }).Count
            } | Measure-Object -Maximum).Maximum
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Remove = [PSCustomObject]@{
            File = @(1..4).ForEach({
                    [PSCustomObject]@{
                        ComputerName = 'localhost'
                        Path         = "c:\file$_.txt"
                        OlderThan    = @{ Quantity = 1; Unit = 'Day' }
                    }
                })
        }
    }
    BeforeEach {
        Remove-Item -Path "$testJobLogFolder\*" -Force
    }
    It 'JobsPerComputer limits the jobs running at once on one computer' {
        $testNewInputFile.MaxConcurrent = [PSCustomObject]@{ JobsTotal = 4; JobsPerComputer = 2 }
        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testNewParams

        @(Get-ChildItem -LiteralPath $testJobLogFolder -File) | Should -HaveCount 4
        Get-MaxOverlapHC | Should -Be 2
    }
    It 'JobsTotal limits the jobs running at once over all computers' {
        $testNewInputFile.MaxConcurrent = [PSCustomObject]@{ JobsTotal = 3; JobsPerComputer = 4 }
        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testNewParams

        @(Get-ChildItem -LiteralPath $testJobLogFolder -File) | Should -HaveCount 4
        Get-MaxOverlapHC | Should -Be 3
    }
}
Describe 'create an Excel file' {
    BeforeAll {
        Clear-TestLogFolderHC

        Mock Invoke-Command {
            $testData[0]
        } -ParameterFilter {
            $FilePath -eq $testParams.Path.RemoveFileScript
        }

        Mock Invoke-Command {
            $testData[1]
            $testData[2]
        } -ParameterFilter {
            $FilePath -eq $testParams.Path.RemoveFilesInFolderScript
        }

        Mock Invoke-Command {
            $testData[3]
        } -ParameterFilter {
            $FilePath -eq $testParams.Path.RemoveEmptyFoldersScript
        }

        Test-NewJsonFileHC $testInputFile

        . $testScript @testParams

        $testExcelLogFile = Get-TestExcelFileHC
    }
    It 'in the log folder' {
        $testExcelLogFile | Should -Not -BeNullOrEmpty
    }
    Context "with sheet 'Overview'" {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    DateTime     = Get-Date
                    ComputerName = $testData[0].ComputerName
                    Type         = $testData[0].Type
                    Path         = $testData[0].FullName
                    CreationTime = $testData[0].CreationTime
                    OlderThan    = "$($testInputFile.Remove.File[0].OlderThan.Quantity) $($testInputFile.Remove.File[0].OlderThan.Unit)"
                    Action       = $testData[0].Action
                    Error        = $testData[0].Error
                }
                @{
                    DateTime     = Get-Date
                    ComputerName = $testData[1].ComputerName
                    Type         = $testData[1].Type
                    Path         = $testData[1].FullName
                    CreationTime = $testData[1].CreationTime
                    OlderThan    = "$($testInputFile.Remove.FilesInFolder[0].OlderThan.Quantity) $($testInputFile.Remove.FilesInFolder[0].OlderThan.Unit)"
                    Action       = $testData[1].Action
                    Error        = $testData[1].Error
                }
                @{
                    DateTime     = Get-Date
                    ComputerName = $testData[2].ComputerName
                    Type         = $testData[2].Type
                    Path         = $testData[2].FullName
                    CreationTime = $testData[2].CreationTime
                    OlderThan    = "$($testInputFile.Remove.FilesInFolder[0].OlderThan.Quantity) $($testInputFile.Remove.FilesInFolder[0].OlderThan.Unit)"
                    Action       = $testData[2].Action
                    Error        = $testData[2].Error
                }
                @{
                    DateTime     = Get-Date
                    ComputerName = $testData[3].ComputerName
                    Type         = $testData[3].Type
                    Path         = $testData[3].FullName
                    CreationTime = $testData[3].CreationTime
                    OlderThan    = $null
                    Action       = $testData[3].Action
                    Error        = $testData[3].Error
                }
            )

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Path -eq $testRow.Path
                }
                $actualRow.ComputerName | Should -Be $testRow.ComputerName
                $actualRow.Type | Should -Be $testRow.Type
                $actualRow.DateTime.ToString('yyyyMMdd') |
                Should -Be $testRow.DateTime.ToString('yyyyMMdd')
                $actualRow.CreationTime.ToString('yyyyMMdd HHmmss') |
                Should -Be $testRow.CreationTime.ToString('yyyyMMdd HHmmss')
                $actualRow.OlderThan | Should -Be $testRow.OlderThan
                $actualRow.Action | Should -Be $testRow.Action
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
    Context "with sheet 'Errors'" {
        BeforeAll {
            Clear-TestLogFolderHC

            Mock Invoke-Command {
                throw 'Oops'
            } -ParameterFilter {
                $FilePath -eq $testParams.Path.RemoveFileScript
            }

            . $testScript @testParams

            $testExportedExcelRows = @(
                @{
                    ComputerName = $testInputFile.Remove.File[0].ComputerName
                    Path         = $testInputFile.Remove.File[0].Path
                    Type         = 'RemoveFile'
                    OlderThan    = "$($testInputFile.Remove.File[0].OlderThan.Quantity) $($testInputFile.Remove.File[0].OlderThan.Unit)"
                    Error        = 'Oops'
                }
            )

            $actual = Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName 'Errors'
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            $testRow = $testExportedExcelRows[0]
            $actual.ComputerName | Should -Be $testRow.ComputerName
            $actual.Path | Should -Be $testRow.Path
            $actual.Type | Should -Be $testRow.Type
            $actual.OlderThan | Should -Be $testRow.OlderThan
            $actual.Error | Should -Be $testRow.Error
        }
    }
}
Describe 'Settings.SendMail.When' {
    Context 'send no e-mail when' {
        BeforeAll {
            Mock Invoke-Command
        }
        It "'<_>' and there are no errors and no actions" -ForEach @(
            'Never', 'OnError', 'OnErrorOrAction'
        ) {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Settings.SendMail.When = $_

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should -Not -Invoke Send-MailKitMessageHC -Scope It
        }
    }
    Context 'send an e-mail when' {
        It "'OnError' and there are errors" {
            Mock Invoke-Command { $testData[1] }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Settings.SendMail.When = 'OnError'

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It
        }
        It "'OnErrorOrAction' and there are actions but no errors" {
            Mock Invoke-Command { $testData[0] }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Settings.SendMail.When = 'OnErrorOrAction'

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It
        }
        It "'OnErrorOrAction' and there are errors but no actions" {
            Mock Invoke-Command { $testData[1] }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Settings.SendMail.When = 'OnErrorOrAction'

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It
        }
    }
}
Describe 'send an e-mail' {
    BeforeAll {
        Clear-TestLogFolderHC

        Mock Invoke-Command {
            $testData[0]
            $testData[1]
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Remove = [PSCustomObject]@{
            File = $testNewInputFile.Remove.File
        }

        Test-NewJsonFileHC $testNewInputFile

        . $testScript @testParams
    }
    It 'with the Settings.SendMail properties' {
        Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.Settings.SendMail.To) -and
            ($Bcc -eq $testInputFile.Settings.SendMail.Bcc) -and
            ($From -eq $testInputFile.Settings.SendMail.From) -and
            ($SmtpServerName -eq $testInputFile.Settings.SendMail.Smtp.ServerName) -and
            ($SmtpPort -eq $testInputFile.Settings.SendMail.Smtp.Port) -and
            ($SmtpConnectionType -eq $testInputFile.Settings.SendMail.Smtp.ConnectionType) -and
            ($Credential.UserName -eq $testInputFile.Settings.SendMail.Smtp.UserName) -and
            ($MailKitAssemblyPath -eq $testInputFile.Settings.SendMail.AssemblyPath.MailKit) -and
            ($MimeKitAssemblyPath -eq $testInputFile.Settings.SendMail.AssemblyPath.MimeKit)
        }
    }
    It 'with the correct subject, priority and attachments' {
        Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($Priority -eq 'High') -and
            ($Subject -eq '1 removed, 1 error') -and
            ($Attachments -like '*Log.xlsx') -and
            ($Attachments -like '* - System errors log.json')
        }
    }
    It 'with the correct body' {
        Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($Body -like '*Email body*') -and
            ($Body -like (
                "*<a href=`"{0}`">{1}</a><br>Remove file older than 1 day<br>Removed: 1, <b style=`"color:red;`">errors: 1*" -f $(
                    "\\$($testNewInputFile.Remove.File[0].ComputerName)\z$\$($testNewInputFile.Remove.File[0].Path.Substring(3))"
                ),
                $(
                    $testNewInputFile.Remove.File[0].Name
                )
            ))
        }
    }
    It 'with Settings.SendMail.Subject added to the subject' {
        $testNewInputFile.Settings.SendMail.Subject = 'Custom'
        Test-NewJsonFileHC $testNewInputFile

        . $testScript @testParams

        Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope It -ParameterFilter {
            $Subject -eq '1 removed, 1 error, Custom'
        }
    }
}
Describe 'Settings.SaveInEventLog' {
    BeforeAll {
        Mock Invoke-Command { throw 'Oops' }
    }
    It 'writes job errors to the event log when Save is true' {
        Test-NewJsonFileHC $testInputFile

        .$testScript @testParams

        Should -Invoke Write-EventLog -Scope It -ParameterFilter {
            ($LogName -eq $testInputFile.Settings.SaveInEventLog.LogName) -and
            ($Source -eq $testInputFile.Settings.ScriptName) -and
            ($EntryType -eq 'Error') -and
            ($Message -like '*Oops*')
        }
    }
    It 'writes nothing to the event log when Save is false' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SaveInEventLog.Save = $false

        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testParams

        Should -Not -Invoke Write-EventLog -Scope It
    }
}
Describe 'Settings.SaveLogFiles.DeleteLogsAfterDays' {
    It 'removes log files older than the given days' {
        $testOldLogFile = New-Item -Path "$testLogFolder\old log.txt" -ItemType File -Force
        $testOldLogFile.LastWriteTime = (Get-Date).AddDays(-3)

        $testNewLogFile = New-Item -Path "$testLogFolder\new log.txt" -ItemType File -Force

        Test-NewJsonFileHC $testInputFile

        .$testScript @testParams

        $testOldLogFile.FullName | Should -Not -Exist
        $testNewLogFile.FullName | Should -Exist
    }
}
