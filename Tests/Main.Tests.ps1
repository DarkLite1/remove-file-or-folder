#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.2.0' }
#Requires -Modules ImportExcel

BeforeAll {
    $testLogFolder = (New-Item 'TestDrive:/log' -ItemType Directory).FullName

    $testInputFile = @{
        MaxConcurrent = @{
            JobsTotal       = 1
            JobsPerComputer = 1
        }
        Tasks         = @(
            @{
                ComputerName = 'PC1'
                Files        = @(
                    @{ Name = 'FTP log file'; Path = 'z:\file.txt' }
                )
                OlderThan    = @{
                    Quantity = 1
                    Unit     = 'Day'
                    BasedOn  = 'CreationTime'
                }
            }
            @{
                ComputerName       = 'PC2'
                Folders            = @(
                    @{ Name = 'App log folder'; Path = 'z:\folder' }
                )
                OlderThan          = @{
                    Quantity = 1
                    Unit     = 'Day'
                    BasedOn  = 'LastWriteTime'
                }
                Recurse            = $true
                RemoveEmptyFolders = $false
            }
            @{
                ComputerName       = 'PC3'
                Folders            = @(
                    @{ Name = 'Delivery notes'; Path = 'z:\folder' }
                )
                RemoveEmptyFolders = $true
            }
        )
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
            ComputerName = $testInputFile.Tasks[0].ComputerName
            Type         = 'File'
            FullName     = 'z:\file1.txt'
            CreationTime = Get-Date
            Action       = 'Removed'
            Error        = $null
        }
        @{
            DateTime     = Get-Date
            ComputerName = $testInputFile.Tasks[1].ComputerName
            Type         = 'File'
            FullName     = 'z:\file2.txt'
            CreationTime = Get-Date
            Action       = $null
            Error        = 'File in use'
        }
        @{
            DateTime     = Get-Date
            ComputerName = $testInputFile.Tasks[1].ComputerName
            Type         = 'File'
            FullName     = 'z:\file3.txt'
            CreationTime = Get-Date
            Action       = 'Removed'
            Error        = $null
        }
        @{
            DateTime     = Get-Date
            ComputerName = $testInputFile.Tasks[2].ComputerName
            Type         = 'EmptyFolder'
            FullName     = 'z:\folder'
            CreationTime = Get-Date
            Action       = 'Removed'
            Error        = $null
        }
    )

    $testScript = Join-Path (Split-Path $PSScriptRoot) (
        (Split-Path $PSCommandPath -Leaf).Replace('.Tests.ps1', '.ps1')
    )
    $testParams = @{
        ConfigurationJsonFile = $testOutParams.FilePath
        RemoveItemsScript     = (New-Item 'TestDrive:/removeItems.ps1' -ItemType File).FullName
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
        Should-BeTrue
    }
}
Describe 'grouped Exclude configuration' {
    BeforeEach {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
    }
    It 'routes grouped exclusions to each worker type' {
        $testNewInputFile.Tasks[0] | Add-Member Exclude @{ Attributes = @('Hidden') }
        $testNewInputFile.Tasks[1].RemoveEmptyFolders = $true
        $testNewInputFile.Tasks[1] | Add-Member Exclude @{
            Attributes = @('Hidden', 'System')
            Folders = @('z:\folder\keep')
            Files = @('z:\folder\state.json')
        }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        Should-Invoke Invoke-Command -Exactly -Times 1 -Scope It -ParameterFilter {
            ($ArgumentList[0] -eq 'File') -and (($ArgumentList[8] -join ',') -eq 'Hidden')
        }
        Should-Invoke Invoke-Command -Exactly -Times 1 -Scope It -ParameterFilter {
            ($ArgumentList[0] -eq 'FilesInFolder') -and
            (($ArgumentList[2] -join ',') -eq 'z:\folder\keep') -and
            (($ArgumentList[6] -join ',') -eq 'z:\folder\state.json') -and
            (($ArgumentList[8] -join ',') -eq 'Hidden,System')
        }
        Should-Invoke Invoke-Command -Exactly -Times 1 -Scope It -ParameterFilter {
            ($ArgumentList[0] -eq 'EmptyFolders') -and ($ArgumentList[1] -eq 'z:\folder') -and
            (($ArgumentList[2] -join ',') -eq 'z:\folder\keep') -and
            (($ArgumentList[8] -join ',') -eq 'Hidden,System')
        }
    }
    It 'accepts empty grouped settings <Label>' -ForEach @(
        @{ Label = 'object'; Value = @{} }
        @{ Label = 'arrays'; Value = @{ Attributes = @(); Folders = @(); Files = @() } }
    ) {
        $testNewInputFile.Tasks[1] | Add-Member Exclude $Value
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        Should-Invoke Invoke-Command -Exactly -Times 1 -Scope It -ParameterFilter {
            ($ArgumentList[0] -eq 'FilesInFolder') -and
            ($ArgumentList[2].Count -eq 0) -and ($ArgumentList[6].Count -eq 0) -and ($ArgumentList[8].Count -eq 0)
        }
    }
    It 'rejects invalid grouped settings <Label>' -ForEach @(
        @{ Label = 'null'; Value = $null; Property = 'Exclude' }
        @{ Label = 'array'; Value = @(); Property = 'Exclude' }
        @{ Label = 'string'; Value = 'Hidden'; Property = 'Exclude' }
        @{ Label = 'unknown key'; Value = @{ Attribute = @('Hidden') }; Property = 'Exclude.Attribute' }
        @{ Label = 'invalid attribute'; Value = @{ Attributes = @('ReadOnly') }; Property = 'Exclude.Attributes' }
        @{ Label = 'scalar attributes'; Value = @{ Attributes = 'Hidden' }; Property = 'Exclude.Attributes' }
        @{ Label = 'null folders'; Value = @{ Folders = $null }; Property = 'Exclude.Folders' }
        @{ Label = 'scalar files'; Value = @{ Files = 'z:\folder\keep.txt' }; Property = 'Exclude.Files' }
        @{ Label = 'outside folder'; Value = @{ Folders = @('z:\other') }; Property = 'Exclude.Folders' }
        @{ Label = 'outside file'; Value = @{ Files = @('z:\other\keep.txt') }; Property = 'Exclude.Files' }
    ) {
        $testNewInputFile.Tasks[1] | Add-Member Exclude $Value
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        ((Get-TestSystemErrorsHC).Message -join "`n") | Should-BeLikeString "*$([WildcardPattern]::Escape("Tasks[1].$Property"))*"
        Should-Invoke Invoke-Command -Times 0 -Exactly -Scope It
    }
    It 'rejects unsupported flat <_>' -ForEach @('ExcludeAttributes', 'ExcludeFolders', 'ExcludeFiles') {
        $testNewInputFile.Tasks[1] | Add-Member $_ @()
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        ((Get-TestSystemErrorsHC).Message -join "`n") | Should-BeLikeString "*$([WildcardPattern]::Escape("Tasks[1].$_"))*is not supported*"
        Should-Invoke Invoke-Command -Times 0 -Exactly -Scope It
    }
}
Describe 'IgnoreMissingPath configuration' {
    BeforeEach {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $global:LASTEXITCODE = 0
    }
    It 'passes <Label> to every worker type' -ForEach @(
        @{ Label = 'enabled'; Value = $true }
        @{ Label = 'disabled'; Value = $false }
        @{ Label = 'omitted'; Value = $null }
    ) {
        if ($null -ne $Value) {
            foreach ($task in $testNewInputFile.Tasks) {
                $task | Add-Member IgnoreMissingPath $Value
            }
        }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        foreach ($type in 'File', 'FilesInFolder', 'EmptyFolders') {
            Should-Invoke Invoke-Command -Exactly -Times 1 -Scope It -ParameterFilter {
                ($ArgumentList[0] -eq $type) -and ($ArgumentList[9] -ceq [bool]$Value)
            }
        }
    }
    It 'rejects <Label> before running any jobs' -ForEach @(
        @{ Label = 'null'; Value = $null }
        @{ Label = 'a string'; Value = 'true' }
        @{ Label = 'a number'; Value = 1 }
        @{ Label = 'an array'; Value = @($true) }
        @{ Label = 'an object'; Value = @{ Enabled = $true } }
    ) {
        $testNewInputFile.Tasks[0] | Add-Member IgnoreMissingPath $Value
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 1
        (Get-TestSystemErrorsHC).Message | Should-BeLikeString '*Tasks[[]0].IgnoreMissingPath*needs a boolean*'
        Should-NotInvoke Invoke-Command -Scope It
        Should-NotInvoke New-PSSession -Scope It
    }
    It 'reports no errors for optional missing files and folders with JobsTotal <JobsTotal>' -ForEach @(
        @{ JobsTotal = 1 }
        @{ JobsTotal = 3 }
    ) {
        $testRoot = (New-Item "TestDrive:/optional-$JobsTotal" -ItemType Directory).FullName
        $testNewInputFile.MaxConcurrent.JobsTotal = $JobsTotal
        $testNewInputFile.Tasks = @(
            [pscustomobject]@{
                ComputerName = 'localhost'
                Files = @((Join-Path $testRoot 'missing.txt'))
                IgnoreMissingPath = $true
                OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'CreationTime' }
            }
            [pscustomobject]@{
                ComputerName = 'localhost'
                Folders = @((Join-Path $testRoot 'missing-folder'))
                IgnoreMissingPath = $true
                OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'CreationTime' }
                Recurse = $true
                RemoveEmptyFolders = $true
            }
        )
        Test-NewJsonFileHC $testNewInputFile

        & $testScript -ConfigurationJsonFile $testOutParams.FilePath

        $LASTEXITCODE | Should-Be 0
        $testHtmlFile = @(Get-ChildItem -LiteralPath $testLogFolder -Filter '* - Mail.html')
        $testHtmlFile | Should-BeCollection -Count 1
        $html = Get-Content -LiteralPath $testHtmlFile[0].FullName -Raw
        $html | Should-NotBeLikeString '*Path not found*'
        Should-Invoke Send-MailKitMessageHC -Exactly -Times 1 -Scope It -ParameterFilter {
            ($Subject -eq '0 removed') -and ($Body -ceq $html)
        }
        Should-NotInvoke Write-EventLog -Scope It -ParameterFilter { $EntryType -eq 'Error' }
    }
}
Describe 'Exclude.Attributes configuration' {
    BeforeEach {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
    }
    It 'passes attributes to all worker types' {
        foreach ($task in $testNewInputFile.Tasks) {
            $task | Add-Member Exclude @{ Attributes = @('Hidden', 'System') }
        }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        foreach ($type in 'File', 'FilesInFolder', 'EmptyFolders') {
            Should-Invoke Invoke-Command -Exactly -Times 1 -Scope It -ParameterFilter {
                ($ArgumentList[0] -eq $type) -and (($ArgumentList[8] -join ',') -eq 'Hidden,System')
            }
        }
    }
    It 'protects matching items with the real worker and JobsTotal <JobsTotal>' -ForEach @(
        @{ JobsTotal = 1 }
        @{ JobsTotal = 3 }
    ) {
        $testRoot = (New-Item "TestDrive:/main-attributes-$JobsTotal" -ItemType Directory).FullName
        $testProtected = New-Item "$testRoot/system" -ItemType Directory
        $testInside = New-Item "$testRoot/system/keep.txt" -ItemType File
        $testHidden = New-Item "$testRoot/hidden.txt" -ItemType File
        $testExcludedFolder = New-Item "$testRoot/keep" -ItemType Directory
        $testExcludedInside = New-Item "$testRoot/keep/inside.txt" -ItemType File
        $testExcludedFile = New-Item "$testRoot/state.json" -ItemType File
        $testRemove = New-Item "$testRoot/remove.txt" -ItemType File
        $testEmpty = New-Item "$testRoot/empty" -ItemType Directory
        $testProtected.Attributes = $testProtected.Attributes -bor [System.IO.FileAttributes]::System
        $testHidden.Attributes = $testHidden.Attributes -bor [System.IO.FileAttributes]::Hidden
        $testNewInputFile.MaxConcurrent.JobsTotal = $JobsTotal
        $testNewInputFile.Tasks = @([pscustomobject]@{
            ComputerName = 'localhost'
            Folders = @($testRoot)
            OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'LastWriteTime' }
            Recurse = $true
            RemoveEmptyFolders = $true
        })
        $testExclusions = @{
            Attributes = @('Hidden', 'System')
            Folders = @($testExcludedFolder.FullName)
            Files = @($testExcludedFile.FullName)
        }
        $testNewInputFile.Tasks[0] | Add-Member Exclude $testExclusions
        Test-NewJsonFileHC $testNewInputFile
        $global:LASTEXITCODE = 0

        & $testScript -ConfigurationJsonFile $testOutParams.FilePath

        $LASTEXITCODE | Should-Be 0
        (Test-Path -LiteralPath $testInside.FullName) | Should-BeTrue
        (Test-Path -LiteralPath $testHidden.FullName) | Should-BeTrue
        (Test-Path -LiteralPath $testExcludedInside.FullName) | Should-BeTrue
        (Test-Path -LiteralPath $testExcludedFile.FullName) | Should-BeTrue
        (Test-Path -LiteralPath $testRemove.FullName) | Should-BeFalse
        (Test-Path -LiteralPath $testEmpty.FullName) | Should-BeFalse
        $rows = @(Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName Overview)
        $rows | Should-BeCollection -Count 2
        @($rows | Where-Object Error) | Should-BeCollection -Count 0
        Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It -ParameterFilter {
            ($Subject -eq '2 removed') -and ($Body -like '*excluding 1 folder and 1 file and hidden/system items*')
        }
    }
    It 'accepts an empty list and defaults omitted settings to no exclusions' {
        $testNewInputFile.Tasks[0] | Add-Member Exclude @{ Attributes = @() }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        Should-Invoke Invoke-Command -Exactly -Times 3 -Scope It -ParameterFilter {
            @($ArgumentList[8]).Count -eq 0
        }
    }
    It 'rejects <Label>' -ForEach @(
        @{ Label = 'a string'; Value = 'Hidden' }
        @{ Label = 'null'; Value = $null }
        @{ Label = 'an unsupported attribute'; Value = @('ReadOnly') }
        @{ Label = 'a numeric attribute'; Value = @(2) }
        @{ Label = 'an empty attribute'; Value = @('') }
    ) {
        $testNewInputFile.Tasks[0] | Add-Member Exclude @{ Attributes = $Value }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 1
        ((Get-TestSystemErrorsHC).Message -join "`n") | Should-MatchString 'Tasks\[0\]\.Exclude\.Attributes'
        Should-NotInvoke Invoke-Command -Scope It
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

        $LASTEXITCODE | Should-Be 1
        ($testWarnings -join "`n") | Should-BeLikeString '*nonExisting.json*'
        Should-NotInvoke Invoke-Command -Scope It
    }
    It 'RemoveItemsScript not found' {
        Test-NewJsonFileHC $testInputFile

        $testNewParams = $testParams.Clone()
        $testNewParams.RemoveItemsScript = 'c:\NotExisting.ps1'

        .$testScript @testNewParams

        $LASTEXITCODE | Should-Be 1
        ((Get-TestSystemErrorsHC).Message -join "`n") |
        Should-BeLikeString "*RemoveItemsScript 'c:\NotExisting.ps1' not found*"
        Should-NotInvoke Invoke-Command -Scope It
    }
    It '<Description>' -ForEach @(
        @{
            Description = 'MaxConcurrent missing'
            Change      = { param($f) $f.PSObject.Properties.Remove('MaxConcurrent') }
            Message     = "Property 'MaxConcurrent' not found"
        }
        @{
            Description = 'Tasks missing'
            Change      = { param($f) $f.PSObject.Properties.Remove('Tasks') }
            Message     = "Property 'Tasks' not found"
        }
        @{
            Description = 'Tasks is empty'
            Change      = { param($f) $f.Tasks = @() }
            Message     = "Property 'Tasks' not found"
        }
        @{
            Description = 'Remove from the old format is used'
            Change      = { param($f) $f | Add-Member -NotePropertyName 'Remove' -NotePropertyValue @{} }
            Message     = "Property 'Remove' is no longer supported, use 'Tasks' instead."
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
            Message     = "Property 'Settings.SendMail.When' with value 'Sometimes' is not supported"
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
            Message     = "Property 'Settings.SendMail.Smtp.Port' with value '26' is not supported"
        }
        @{
            Description = 'Settings.SendMail.Smtp.ConnectionType not supported'
            Change      = { param($f) $f.Settings.SendMail.Smtp.ConnectionType = 'Wrong' }
            Message     = "Property 'Settings.SendMail.Smtp.ConnectionType' with value 'Wrong' is not supported"
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
            Description = 'Tasks[0] has an unknown property'
            Change      = { param($f) $f.Tasks[0] | Add-Member -NotePropertyName 'Recursive' -NotePropertyValue $true }
            Message     = "Property 'Tasks[0].Recursive' is not supported"
        }
        @{
            Description = 'Tasks[0] has Files and Folders'
            Change      = { param($f) $f.Tasks[0] | Add-Member -NotePropertyName 'Folders' -NotePropertyValue @('z:\a') }
            Message     = "Property 'Tasks[0].Files' and 'Tasks[0].Folders' cannot be used at the same time"
        }
        @{
            Description = 'Tasks[0] has no Files or Folders'
            Change      = { param($f) $f.Tasks[0].PSObject.Properties.Remove('Files') }
            Message     = "Property 'Tasks[0].Files' or 'Tasks[0].Folders' not found"
        }
        @{
            Description = 'Tasks[0].OlderThan missing for Files'
            Change      = { param($f) $f.Tasks[0].PSObject.Properties.Remove('OlderThan') }
            Message     = "Property 'Tasks[0].OlderThan' not found"
        }
        @{
            Description = 'Tasks[0].Recurse used with Files'
            Change      = { param($f) $f.Tasks[0] | Add-Member -NotePropertyName 'Recurse' -NotePropertyValue $true }
            Message     = "Property 'Tasks[0].Recurse' can only be used with 'Tasks[0].Folders'"
        }
        @{
            Description = 'Tasks[0].RemoveEmptyFolders used with Files'
            Change      = { param($f) $f.Tasks[0] | Add-Member -NotePropertyName 'RemoveEmptyFolders' -NotePropertyValue $true }
            Message     = "Property 'Tasks[0].RemoveEmptyFolders' can only be used with 'Tasks[0].Folders'"
        }
        @{
            Description = 'Tasks[1].Recurse missing'
            Change      = { param($f) $f.Tasks[1].PSObject.Properties.Remove('Recurse') }
            Message     = "Property 'Tasks[1].Recurse' not found"
        }
        @{
            Description = 'Tasks[1].Recurse not a boolean'
            Change      = { param($f) $f.Tasks[1].Recurse = 'yes' }
            Message     = "Property 'Tasks[1].Recurse' needs to be true or false, the value 'yes' is not supported."
        }
        @{
            Description = 'Tasks[1].RemoveEmptyFolders missing'
            Change      = { param($f) $f.Tasks[1].PSObject.Properties.Remove('RemoveEmptyFolders') }
            Message     = "Property 'Tasks[1].RemoveEmptyFolders' not found"
        }
        @{
            Description = 'Tasks[1].RemoveEmptyFolders not a boolean'
            Change      = { param($f) $f.Tasks[1].RemoveEmptyFolders = 'yes' }
            Message     = "Property 'Tasks[1].RemoveEmptyFolders' needs to be true or false, the value 'yes' is not supported."
        }
        @{
            Description = 'Tasks[2].Recurse used without OlderThan'
            Change      = { param($f) $f.Tasks[2] | Add-Member -NotePropertyName 'Recurse' -NotePropertyValue $true }
            Message     = "Property 'Tasks[2].Recurse' can only be used together with 'Tasks[2].OlderThan'"
        }
        @{
            Description = 'Tasks[2] has nothing to remove'
            Change      = { param($f) $f.Tasks[2].RemoveEmptyFolders = $false }
            Message     = "Property 'Tasks[2].OlderThan' not found. Use 'OlderThan' to remove files, 'RemoveEmptyFolders' to remove empty folders or both."
        }
        @{
            Description = 'Tasks[0].OlderThan.Unit missing'
            Change      = { param($f) $f.Tasks[0].OlderThan.PSObject.Properties.Remove('Unit') }
            Message     = "Property 'Tasks[0].OlderThan.Unit' not found"
        }
        @{
            Description = 'Tasks[0].OlderThan.Unit not supported'
            Change      = { param($f) $f.Tasks[0].OlderThan.Unit = 'notSupported' }
            Message     = "Property 'Tasks[0].OlderThan.Unit' with value 'notSupported' is not supported. Supported values are 'Day', 'Month' or 'Year'."
        }
        @{
            Description = 'Tasks[0].OlderThan.Quantity missing'
            Change      = { param($f) $f.Tasks[0].OlderThan.PSObject.Properties.Remove('Quantity') }
            Message     = "Property 'Tasks[0].OlderThan.Quantity' not found. Use value 0 to remove all files."
        }
        @{
            Description = 'Tasks[0].OlderThan.Quantity not a number'
            Change      = { param($f) $f.Tasks[0].OlderThan.Quantity = 'a' }
            Message     = "Property 'Tasks[0].OlderThan.Quantity' needs to be a positive number, the value 'a' is not supported."
        }
        @{
            Description = 'Tasks[1].OlderThan.Quantity negative'
            Change      = { param($f) $f.Tasks[1].OlderThan.Quantity = -1 }
            Message     = "Property 'Tasks[1].OlderThan.Quantity' needs to be a positive number, the value '-1' is not supported."
        }
        @{
            Description = 'Tasks[0].OlderThan.BasedOn missing'
            Change      = { param($f) $f.Tasks[0].OlderThan.PSObject.Properties.Remove('BasedOn') }
            Message     = "Property 'Tasks[0].OlderThan.BasedOn' not found. Use 'CreationTime' or 'LastWriteTime'."
        }
        @{
            Description = 'Tasks[1].OlderThan.BasedOn not supported'
            Change      = { param($f) $f.Tasks[1].OlderThan.BasedOn = 'LastAccessTime' }
            Message     = "Property 'Tasks[1].OlderThan.BasedOn' with value 'LastAccessTime' is not supported. Supported values are 'CreationTime' or 'LastWriteTime'."
        }
        @{
            Description = 'Tasks[0].OlderThan has an unknown property'
            Change      = { param($f) $f.Tasks[0].OlderThan | Add-Member -NotePropertyName 'Based' -NotePropertyValue 'x' }
            Message     = "Property 'Tasks[0].OlderThan.Based' is not supported"
        }
        @{
            Description = 'Tasks[0].Files[0] without a path'
            Change      = { param($f) $f.Tasks[0].Files[0].Path = $null }
            Message     = "Property 'Tasks[0].Files[0]' needs a path"
        }
        @{
            Description = 'Tasks[1].Folders[0] has an unknown property'
            Change      = { param($f) $f.Tasks[1].Folders[0] | Add-Member -NotePropertyName 'Other' -NotePropertyValue 1 }
            Message     = "Property 'Tasks[1].Folders[0].Other' is not supported"
        }
        @{
            Description = 'Tasks[1] local path without ComputerName'
            Change      = { param($f) $f.Tasks[1].ComputerName = $null }
            Message     = "Property 'Tasks[1].ComputerName' not found, it is required for the local path 'z:\folder'"
        }
        @{
            Description = 'Tasks[0].Exclude.Folders used with Files'
            Change      = { param($f) $f.Tasks[0] | Add-Member Exclude @{ Folders = @('z:\a') } }
            Message     = "Property 'Tasks[0].Exclude.Folders' can only be used with 'Tasks[0].Folders'"
        }
        @{
            Description = 'Tasks[1].Exclude.Folders contains an empty path'
            Change      = { param($f) $f.Tasks[1] | Add-Member Exclude @{ Folders = @('') } }
            Message     = "Property 'Tasks[1].Exclude.Folders' needs to be an array of folder paths, the value '' is not supported."
        }
        @{
            Description = 'Tasks[1].Exclude.Folders contains a number'
            Change      = { param($f) $f.Tasks[1] | Add-Member Exclude @{ Folders = @(5) } }
            Message     = "Property 'Tasks[1].Exclude.Folders' needs to be an array of folder paths, the value '5' is not supported."
        }
        @{
            Description = 'Tasks[1].Exclude.Folders is not below a folder'
            Change      = { param($f) $f.Tasks[1] | Add-Member Exclude @{ Folders = @('z:\other\keep') } }
            Message     = "Property 'Tasks[1].Exclude.Folders' contains 'z:\other\keep', which is not a subfolder of a path in 'Tasks[1].Folders'"
        }
        @{
            Description = 'Tasks[1].Exclude.Folders is the folder itself'
            Change      = { param($f) $f.Tasks[1] | Add-Member Exclude @{ Folders = @('z:\folder') } }
            Message     = "Property 'Tasks[1].Exclude.Folders' contains 'z:\folder', which is not a subfolder of a path in 'Tasks[1].Folders'"
        }
        @{
            Description = 'Tasks[1].Exclude.Folders escapes the folder with dot segments'
            Change      = { param($f) $f.Tasks[1] | Add-Member Exclude @{ Folders = @('z:\folder\..\other') } }
            Message     = "Property 'Tasks[1].Exclude.Folders' contains 'z:\other', which is not a subfolder of a path in 'Tasks[1].Folders'"
        }
        @{
            Description = 'Tasks[0].Exclude.Files used with Files'
            Change      = { param($f) $f.Tasks[0] | Add-Member Exclude @{ Files = @('z:\a.txt') } }
            Message     = "Property 'Tasks[0].Exclude.Files' can only be used with 'Tasks[0].Folders'"
        }
        @{
            Description = 'Tasks[2].Exclude.Files used without OlderThan'
            Change      = { param($f) $f.Tasks[2] | Add-Member Exclude @{ Files = @('z:\folder\a.txt') } }
            Message     = "Property 'Tasks[2].Exclude.Files' can only be used together with 'Tasks[2].OlderThan'"
        }
        @{
            Description = 'Tasks[1].Exclude.Files contains an empty path'
            Change      = { param($f) $f.Tasks[1] | Add-Member Exclude @{ Files = @('') } }
            Message     = "Property 'Tasks[1].Exclude.Files' needs to be an array of file paths, the value '' is not supported."
        }
        @{
            Description = 'Tasks[1].Exclude.Files is not below a folder'
            Change      = { param($f) $f.Tasks[1] | Add-Member Exclude @{ Files = @('z:\other\a.txt') } }
            Message     = "Property 'Tasks[1].Exclude.Files' contains 'z:\other\a.txt', which is not a file of a path in 'Tasks[1].Folders'"
        }
    ) {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        & $Change $testNewInputFile

        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testParams -WarningVariable testWarnings -WarningAction SilentlyContinue

        $testMessage = "*$([WildcardPattern]::Escape($Message))*"

        $LASTEXITCODE | Should-Be 1
        ($testWarnings -join "`n") | Should-BeLikeString $testMessage
        ((Get-TestSystemErrorsHC).Message -join "`n") | Should-BeLikeString $testMessage
        Should-NotInvoke Invoke-Command -Scope It
    }
    It 'is reported by e-mail and in the event log' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Files[0].Path = $null

        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testParams

        Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It -ParameterFilter {
            ($Priority -eq 'High') -and
            ($Body -like "*Property &#39;Tasks``[0``].Files``[0``]&#39; needs a path*")
        }
        Should-Invoke Write-EventLog -Scope It -ParameterFilter {
            ($EntryType -eq 'Error') -and
            ($Message -like "*Property 'Tasks``[0``].Files``[0``]' needs a path*")
        }
    }
    It 'Settings.SendMail properties are not needed when SendMail.When is Never' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SendMail = [PSCustomObject]@{ When = 'Never' }

        Test-NewJsonFileHC $testNewInputFile

        $global:LASTEXITCODE = 0

        .$testScript @testParams

        $LASTEXITCODE | Should-Be 0
        Should-Invoke Invoke-Command -Times 3 -Exactly -Scope It
        Should-NotInvoke Send-MailKitMessageHC -Scope It
    }
}
Describe 'Example.json' {
    It 'is a valid input file' {
        Clear-TestLogFolderHC

        $testExample = Get-Content -LiteralPath (
            Join-Path (Split-Path $PSScriptRoot) 'Example.json'
        ) -Raw | ConvertFrom-Json

        # keep the examples, but avoid real mail, event log and log folder
        $testExample.Settings.SendMail.When = 'Never'
        $testExample.Settings.SaveInEventLog.Save = $false
        $testExample.Settings.SaveLogFiles.Where.Folder = $testLogFolder
        # mocks only work in sequential mode
        $testExample.MaxConcurrent.JobsTotal = 1

        Test-NewJsonFileHC $testExample

        $global:LASTEXITCODE = 0

        .$testScript @testParams -WarningVariable testWarnings -WarningAction SilentlyContinue

        $testWarnings | Should-BeCollection -Count 0
        $LASTEXITCODE | Should-Be 0
        Should-Invoke Invoke-Command -Scope It
    }
}
Describe 'execute script' {
    Context 'Files' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @($testNewInputFile.Tasks[0])

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should-Invoke New-PSSession -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq $testNewInputFile.Tasks[0].ComputerName) -and
                ($ConfigurationName -eq 'PowerShell.7')
            }
            Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Session) -and
                ($FilePath -eq $testParams.RemoveItemsScript) -and
                ($ArgumentList[0] -eq 'File') -and
                ($ArgumentList[1] -eq $testNewInputFile.Tasks[0].Files[0].Path) -and
                ($ArgumentList[3] -eq $testNewInputFile.Tasks[0].OlderThan.Unit) -and
                ($ArgumentList[4] -eq $testNewInputFile.Tasks[0].OlderThan.Quantity) -and
                ($ArgumentList[7] -eq 'CreationTime')
            }
        }
    }
    Context 'Folders with OlderThan' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @($testNewInputFile.Tasks[1])

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should-Invoke New-PSSession -Times 1 -Exactly -Scope Context -ParameterFilter {
                $ComputerName -eq $testNewInputFile.Tasks[0].ComputerName
            }
            Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Session) -and
                ($FilePath -eq $testParams.RemoveItemsScript) -and
                ($ArgumentList[0] -eq 'FilesInFolder') -and
                ($ArgumentList[1] -eq $testNewInputFile.Tasks[0].Folders[0].Path) -and
                ($ArgumentList[3] -eq $testNewInputFile.Tasks[0].OlderThan.Unit) -and
                ($ArgumentList[4] -eq $testNewInputFile.Tasks[0].OlderThan.Quantity) -and
                ($ArgumentList[5] -eq $testNewInputFile.Tasks[0].Recurse) -and
                ($ArgumentList[7] -eq 'LastWriteTime')
            }
        }
    }
    Context 'Folders with RemoveEmptyFolders' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @($testNewInputFile.Tasks[2])

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should-Invoke New-PSSession -Times 1 -Exactly -Scope Context -ParameterFilter {
                $ComputerName -eq $testNewInputFile.Tasks[0].ComputerName
            }
            Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Session) -and
                ($FilePath -eq $testParams.RemoveItemsScript) -and
                ($ArgumentList[0] -eq 'EmptyFolders') -and
                ($ArgumentList[1] -eq $testNewInputFile.Tasks[0].Folders[0].Path)
            }
        }
        It 'and close the session' {
            Should-Invoke Remove-PSSession -Times 1 -Exactly -Scope Context
        }
    }
    Context 'Folders with OlderThan and RemoveEmptyFolders as plain path strings' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @($testNewInputFile.Tasks[1])
            $testNewInputFile.Tasks[0].Folders = @('z:\a', 'z:\b')
            $testNewInputFile.Tasks[0].RemoveEmptyFolders = $true

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'removes the files in every folder' {
            foreach ($testPath in 'z:\a', 'z:\b') {
                Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                    ($ArgumentList[0] -eq 'FilesInFolder') -and
                    ($ArgumentList[1] -eq $testPath)
                }
            }
        }
        It 'removes the empty folders in every folder' {
            foreach ($testPath in 'z:\a', 'z:\b') {
                Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                    ($ArgumentList[0] -eq 'EmptyFolders') -and
                    ($ArgumentList[1] -eq $testPath)
                }
            }
        }
    }
    Context 'Exclude.Folders' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @($testNewInputFile.Tasks[1])
            $testNewInputFile.Tasks[0].Folders = @('z:\a', 'z:\b')
            $testNewInputFile.Tasks[0].RemoveEmptyFolders = $true
            $testNewInputFile.Tasks[0] | Add-Member Exclude @{ Folders = @('z:\a\keep', 'z:\a\also keep') }

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'are passed to the jobs of the folder they are in' {
            foreach ($testType in 'FilesInFolder', 'EmptyFolders') {
                Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                    ($ArgumentList[0] -eq $testType) -and
                    ($ArgumentList[1] -eq 'z:\a') -and
                    (($ArgumentList[2] -join '|') -eq 'z:\a\keep|z:\a\also keep')
                }
            }
        }
        It 'are not passed to the jobs of other folders' {
            foreach ($testType in 'FilesInFolder', 'EmptyFolders') {
                Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                    ($ArgumentList[0] -eq $testType) -and
                    ($ArgumentList[1] -eq 'z:\b') -and
                    (@($ArgumentList[2]).Count -eq 0)
                }
            }
        }
    }
    Context 'Exclude.Files' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @($testNewInputFile.Tasks[1])
            $testNewInputFile.Tasks[0].Folders = @('z:\a', 'z:\b')
            $testNewInputFile.Tasks[0] | Add-Member Exclude @{ Files = @('z:\a\keep.txt', 'z:\a\sub\keep.json') }

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams
        }
        It 'are passed to the job of the folder they are in' {
            Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ArgumentList[0] -eq 'FilesInFolder') -and
                ($ArgumentList[1] -eq 'z:\a') -and
                (($ArgumentList[6] -join '|') -eq 'z:\a\keep.txt|z:\a\sub\keep.json')
            }
        }
        It 'are not passed to the jobs of other folders' {
            Should-Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ArgumentList[0] -eq 'FilesInFolder') -and
                ($ArgumentList[1] -eq 'z:\b') -and
                (@($ArgumentList[6]).Count -eq 0)
            }
        }
    }
    Context 'normalized exclusion routing' {
        It 'normalizes roots and routes dot-segment exclusions to the correct job' {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @($testNewInputFile.Tasks[1])
            $testNewInputFile.Tasks[0].Folders = @('z:/a/.', 'z:/b')
            $testNewInputFile.Tasks[0] | Add-Member Exclude @{
                Folders = @('z:/b/../a/keep')
                Files = @('z:/a/sub/../state.json')
            }

            Test-NewJsonFileHC $testNewInputFile
            .$testScript @testParams

            Should-Invoke Invoke-Command -Times 1 -Exactly -Scope It -ParameterFilter {
                ($ArgumentList[0] -eq 'FilesInFolder') -and
                ($ArgumentList[1].TrimEnd('\') -eq 'z:\a') -and
                (($ArgumentList[2] -join '|') -eq 'z:\a\keep') -and
                (($ArgumentList[6] -join '|') -eq 'z:\a\state.json')
            }
            Should-Invoke Invoke-Command -Times 1 -Exactly -Scope It -ParameterFilter {
                ($ArgumentList[1] -eq 'z:\b') -and
                (@($ArgumentList[2]).Count -eq 0) -and
                (@($ArgumentList[6]).Count -eq 0)
            }
        }
    }
    Context 'PSSessionConfiguration' {
        It 'is used for the remote session' {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @($testNewInputFile.Tasks[0])
            $testNewInputFile | Add-Member -NotePropertyName 'PSSessionConfiguration' -NotePropertyValue 'PowerShell.7.5'

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should-Invoke New-PSSession -Times 1 -Exactly -Scope It -ParameterFilter {
                $ConfigurationName -eq 'PowerShell.7.5'
            }
        }
    }
}
Describe 'retry a remote job' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @($testNewInputFile.Tasks[0])

        Test-NewJsonFileHC $testNewInputFile

        $testTransientError = 'Processing data from remote server PC1 failed: The I/O operation has been aborted because of either a thread exit or an application request.'
    }
    BeforeEach {
        Clear-TestLogFolderHC
    }
    It 'up to 3 times on a transient WinRM abort and report the last error' {
        Mock Invoke-Command { throw $testTransientError }

        .$testScript @testParams

        Should-Invoke Invoke-Command -Times 3 -Exactly
        Should-Invoke New-PSSession -Times 3 -Exactly
        Should-Invoke Remove-PSSession -Times 3 -Exactly
        Should-Invoke Start-Sleep -Times 2 -Exactly

        $actual = Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName 'Errors'
        $actual.Error | Should-BeLikeString '*I/O operation has been aborted*'
    }
    It 'and keep the result when a retry succeeds' {
        $script:testAttempt = 0
        Mock Invoke-Command {
            $script:testAttempt++
            if ($script:testAttempt -eq 1) { throw $testTransientError }
            $testData[0]
        }

        .$testScript @testParams

        Should-Invoke Invoke-Command -Times 2 -Exactly

        $testExcelFile = Get-TestExcelFileHC
        $actual = Import-Excel -Path $testExcelFile.FullName -WorksheetName 'Overview'
        $actual.Path | Should-Be $testData[0].FullName
        Get-ExcelSheetInfo -Path $testExcelFile.FullName |
        Where-Object Name -EQ 'Errors' | Should-BeCollection -Count 0
    }
    It 'shows attempt, path and delay messages when verbose is enabled' {
        $script:testAttempt = 0
        Mock Invoke-Command {
            $script:testAttempt++
            if ($script:testAttempt -eq 1) { throw $testTransientError }
            $testData[0]
        }

        $actual = @(& $testScript @testParams -Verbose 4>&1)
        $messages = @($actual | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] } | ForEach-Object Message)
        $messages | Should-ContainCollection "Starting job 'RemoveFile' on 'PC1' for 'z:\file.txt' (attempt 1 of 3)"
        $messages | Should-ContainCollection "Retrying job 'RemoveFile' on 'PC1' for 'z:\file.txt' after WinRM abort; attempt 2 of 3 in 5 seconds"
        $messages | Should-ContainCollection "Starting job 'RemoveFile' on 'PC1' for 'z:\file.txt' (attempt 2 of 3)"
        $messages | Should-ContainCollection 'Run summary: 1 removed, 0 errors'
        Should-Invoke Invoke-Command -Times 2 -Exactly -ParameterFilter { $PesterBoundParameters.Verbose -eq $true }
    }
    It 'not on other errors' {
        Mock Invoke-Command { throw 'Oops' }

        .$testScript @testParams

        Should-Invoke Invoke-Command -Times 1 -Exactly
        Should-Invoke Remove-PSSession -Times 1 -Exactly
        Should-NotInvoke Start-Sleep
    }
    It 'shows final error counts and a direct failure warning without preparation text' {
        Mock Invoke-Command { throw 'Permanent failure' }
        $testWarnings = @()

        $actual = @(& $testScript @testParams -Verbose -WarningVariable testWarnings -WarningAction SilentlyContinue 4>&1)
        $messages = @($actual | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] } | ForEach-Object Message)

        $messages | Should-ContainCollection 'Run summary: 0 removed, 1 errors'
        ($testWarnings -join ' ') | Should-BeLikeString "*Job 'RemoveFile' failed on 'PC1' for 'z:\file.txt': Permanent failure*"
        @($messages | Where-Object { $_ -like 'Retrying job*' }) | Should-BeCollection -Count 0
    }
}
Describe 'main diagnostic messages with the real worker' {
    It 'respects Verbose <VerboseEnabled> with JobsTotal <JobsTotal> without contaminating results' -ForEach @(
        foreach ($jobsTotal in 1, 3) {
            foreach ($verboseEnabled in $false, $true) {
                @{ JobsTotal = $jobsTotal; VerboseEnabled = $verboseEnabled }
            }
        }
    ) {
        Clear-TestLogFolderHC
        $testRoot = (New-Item "TestDrive:/messages-$JobsTotal-$VerboseEnabled" -ItemType Directory).FullName
        $testFile = New-Item "$testRoot/file.txt" -ItemType File
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SendMail.When = 'Never'
        $testNewInputFile.Settings.SaveInEventLog.Save = $false
        $testNewInputFile.MaxConcurrent.JobsTotal = $JobsTotal
        $testNewInputFile.Tasks = @([pscustomobject]@{
                ComputerName = 'localhost'
                Files = @($testFile.FullName)
                OlderThan = @{ Unit = 'Day'; Quantity = 0; BasedOn = 'CreationTime' }
            })
        Test-NewJsonFileHC $testNewInputFile
        $global:LASTEXITCODE = 0

        $actual = @(& $testScript -ConfigurationJsonFile $testOutParams.FilePath -Verbose:$VerboseEnabled 4>&1)
        $messages = @($actual | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] } | ForEach-Object Message)

        $LASTEXITCODE | Should-Be 0
        (Test-Path -LiteralPath $testFile.FullName) | Should-BeFalse
        $rows = @(Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName Overview)
        $rows | Should-BeCollection -Count 1
        $rows[0].Action | Should-Be Removed
        $rows[0].Path | Should-Be $testFile.FullName
        if ($VerboseEnabled) {
            @($messages | Where-Object { $_ -like "Prepared job 'RemoveFile'*" }) | Should-BeCollection -Count 1
            $startMessage = "Starting job 'RemoveFile' on '$env:COMPUTERNAME' for '$($testFile.FullName)' (attempt 1 of 3)"
            $messages | Should-ContainCollection $startMessage
            $messages | Should-ContainCollection "Removed file '$($testFile.FullName)'"
            $messages | Should-ContainCollection 'Run summary: 1 removed, 0 errors'
            [array]::IndexOf($messages, $startMessage) | Should-BeLessThan ([array]::IndexOf($messages, "Removed file '$($testFile.FullName)'"))
        }
        else { $messages | Should-BeCollection -Count 0 }
    }
}
Describe 'MaxConcurrent' {
    BeforeAll {
        $testJobLogFolder = (New-Item 'TestDrive:/jobLog' -ItemType Directory).FullName

        $testJobScript = (New-Item 'TestDrive:/job.ps1' -ItemType File).FullName
        Set-Content -LiteralPath $testJobScript -Value @"
param(`$Type, `$Path, `$Unit, `$Quantity)
`$start = [DateTime]::UtcNow.Ticks
Start-Sleep -Milliseconds 1500
`$end = [DateTime]::UtcNow.Ticks
Set-Content -LiteralPath (Join-Path '$testJobLogFolder' ([guid]::NewGuid())) -Value "`$start;`$end"
"@

        $testNewParams = $testParams.Clone()
        $testNewParams.RemoveItemsScript = $testJobScript

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
        $testNewInputFile.Tasks = @(
            [PSCustomObject]@{
                ComputerName = 'localhost'
                Files        = @(1..4).ForEach({ "c:\file$_.txt" })
                OlderThan    = @{ Quantity = 1; Unit = 'Day'; BasedOn = 'CreationTime' }
            }
        )
    }
    BeforeEach {
        Remove-Item -Path "$testJobLogFolder\*" -Force
    }
    It 'JobsPerComputer limits the jobs running at once on one computer' {
        $testNewInputFile.MaxConcurrent = [PSCustomObject]@{ JobsTotal = 4; JobsPerComputer = 2 }
        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testNewParams

        @(Get-ChildItem -LiteralPath $testJobLogFolder -File) | Should-BeCollection -Count 4
        Get-MaxOverlapHC | Should-Be 2
    }
    It 'JobsTotal limits the jobs running at once over all computers' {
        $testNewInputFile.MaxConcurrent = [PSCustomObject]@{ JobsTotal = 3; JobsPerComputer = 4 }
        Test-NewJsonFileHC $testNewInputFile

        .$testScript @testNewParams

        @(Get-ChildItem -LiteralPath $testJobLogFolder -File) | Should-BeCollection -Count 4
        Get-MaxOverlapHC | Should-Be 3
    }
}
Describe 'with the real Remove items script on the local computer' {
    BeforeAll {
        Clear-TestLogFolderHC

        $testRoot = (New-Item 'TestDrive:/e2e' -ItemType Directory).FullName
        $testSingleFile = New-Item "$testRoot\single.txt" -ItemType File
        $testFolderFile = New-Item "$testRoot\folder\sub\file.txt" -ItemType File -Force
        $testNewFile = New-Item "$testRoot\folder\new.txt" -ItemType File
        $testOldFile = New-Item "$testRoot\folder\old.txt" -ItemType File
        $testOldFile.CreationTime = (Get-Date).AddDays(-10)
        $testKeepFile = New-Item "$testRoot\folder\keep\PrintHistory.json" -ItemType File -Force
        $testKeepFile.CreationTime = (Get-Date).AddDays(-10)
        $testKeepEmptyFolder = New-Item "$testRoot\folder\keep\empty" -ItemType Directory
        $testKeepSingleFile = New-Item "$testRoot\folder\state.json" -ItemType File
        $testKeepSingleFile.CreationTime = (Get-Date).AddDays(-10)

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @(
            [PSCustomObject]@{
                ComputerName = 'localhost'
                Files        = @($testSingleFile.FullName)
                OlderThan    = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'CreationTime' }
            }
            [PSCustomObject]@{
                ComputerName       = 'localhost'
                Folders            = @("$testRoot\folder")
                Exclude            = @{
                    Folders = @("$testRoot\folder\keep")
                    Files = @("$testRoot\folder\state.json")
                }
                OlderThan          = @{ Quantity = 5; Unit = 'Day'; BasedOn = 'CreationTime' }
                Recurse            = $true
                RemoveEmptyFolders = $true
            }
        )
        $testFolderFile.CreationTime = (Get-Date).AddDays(-10)

        Test-NewJsonFileHC $testNewInputFile

        $testNewParams = $testParams.Clone()
        $testNewParams.Remove('RemoveItemsScript')

        $global:LASTEXITCODE = 0

        . $testScript @testNewParams
    }
    It 'removes the file' {
        (Test-Path -LiteralPath $testSingleFile.FullName) | Should-BeFalse
    }
    It 'removes the old files in the folder and its subfolders' {
        (Test-Path -LiteralPath $testOldFile.FullName) | Should-BeFalse
        (Test-Path -LiteralPath $testFolderFile.FullName) | Should-BeFalse
    }
    It 'keeps the new files' {
        (Test-Path -LiteralPath $testNewFile.FullName) | Should-BeTrue
    }
    It 'keeps the old files and empty folders in an excluded folder' {
        (Test-Path -LiteralPath $testKeepFile.FullName) | Should-BeTrue
        (Test-Path -LiteralPath $testKeepEmptyFolder.FullName) | Should-BeTrue
    }
    It 'keeps an excluded old file' {
        (Test-Path -LiteralPath $testKeepSingleFile.FullName) | Should-BeTrue
    }
    It 'removes the folders that became empty' {
        (Test-Path -LiteralPath "$testRoot\folder\sub") | Should-BeFalse
    }
    It 'exports every removed item to Excel' {
        $actual = Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName 'Overview'

        $actual | Should-BeCollection -Count 4
        $actual.Action | Sort-Object -Unique | Should-Be 'Removed'
    }
    It 'exits without error' {
        $LASTEXITCODE | Should-Be 0
    }
}
Describe 'end-to-end filesystem scenarios' {
    BeforeAll {
        function New-E2EFixtureHC {
            Clear-TestLogFolderHC
            $script:e2eRoot = (New-Item "TestDrive:/scenario-$([guid]::NewGuid())" -ItemType Directory).FullName
            $script:e2eRemovedFiles = [System.Collections.Generic.List[string]]::new()
            $script:e2eRemovedFolders = [System.Collections.Generic.List[string]]::new()
            $script:e2eKeptFiles = @{}
            $script:e2eConfiguration = Copy-ObjectHC $testInputFile
            $script:e2eConfiguration.Settings.SendMail.When = 'Never'
            $script:e2eConfiguration.Settings.SaveInEventLog.Save = $false
            $script:e2eConfiguration.MaxConcurrent.JobsPerComputer = 2
        }

        function Add-E2EFolderHC {
            param ([string]$RelativePath, [switch]$Removed)

            $folderPath = Join-Path $e2eRoot $RelativePath
            $null = New-Item -Path $folderPath -ItemType Directory -Force
            if ($Removed) { $e2eRemovedFolders.Add($folderPath) }
            $folderPath
        }

        function Add-E2EFileHC {
            param (
                [string]$RelativePath,
                [datetime]$CreationTime = (Get-Date),
                [datetime]$LastWriteTime = (Get-Date),
                [switch]$Removed,
                [switch]$Hidden,
                [switch]$ReadOnly
            )

            $filePath = Join-Path $e2eRoot $RelativePath
            $null = New-Item -Path (Split-Path $filePath) -ItemType Directory -Force
            $content = "Keep exact contents: $RelativePath"
            [System.IO.File]::WriteAllText($filePath, $content)
            [System.IO.File]::SetCreationTime($filePath, $CreationTime)
            [System.IO.File]::SetLastWriteTime($filePath, $LastWriteTime)
            $attributes = [System.IO.File]::GetAttributes($filePath)
            if ($Hidden) { $attributes = $attributes -bor [System.IO.FileAttributes]::Hidden }
            if ($ReadOnly) { $attributes = $attributes -bor [System.IO.FileAttributes]::ReadOnly }
            [System.IO.File]::SetAttributes($filePath, $attributes)
            if ($Removed) { $e2eRemovedFiles.Add($filePath) }
            else { $e2eKeptFiles[$filePath] = $content }
            $filePath
        }

        function Invoke-E2EAndAssertHC {
            param ([string[]]$ExpectedErrorPaths = @())

            $expectedFolders = @(
                Get-ChildItem -LiteralPath $e2eRoot -Directory -Recurse -Force |
                Where-Object FullName -NotIn $e2eRemovedFolders |
                Select-Object -ExpandProperty FullName | Sort-Object
            )
            Test-NewJsonFileHC $e2eConfiguration
            $global:LASTEXITCODE = 0
            & $testScript -ConfigurationJsonFile $testOutParams.FilePath
            $LASTEXITCODE | Should-Be $(if ($ExpectedErrorPaths.Count) { 1 } else { 0 })

            (Test-Path -LiteralPath $e2eRoot) | Should-BeTrue
            $actualFiles = @(Get-ChildItem -LiteralPath $e2eRoot -File -Recurse -Force | Select-Object -ExpandProperty FullName | Sort-Object)
            ($actualFiles -join "`n") | Should-Be (($e2eKeptFiles.Keys | Sort-Object) -join "`n")
            foreach ($filePath in $e2eKeptFiles.Keys) {
                [System.IO.File]::ReadAllText($filePath) | Should-BeString -CaseSensitive $e2eKeptFiles[$filePath]
            }
            $actualFolders = @(Get-ChildItem -LiteralPath $e2eRoot -Directory -Recurse -Force | Select-Object -ExpandProperty FullName | Sort-Object)
            ($actualFolders -join "`n") | Should-Be ($expectedFolders -join "`n")
            foreach ($removedPath in @($e2eRemovedFiles) + @($e2eRemovedFolders)) {
                (Test-Path -LiteralPath $removedPath) | Should-BeFalse
            }

            $expectedRemoved = @($e2eRemovedFiles) + @($e2eRemovedFolders)
            if ($expectedRemoved.Count -or $ExpectedErrorPaths.Count) {
                $workbooks = @(Get-TestExcelFileHC)
                $workbooks | Should-BeCollection -Count 1
                $rows = @(Import-Excel -Path $workbooks[0].FullName -WorksheetName Overview)
                $rows | Should-BeCollection -Count ($expectedRemoved.Count + $ExpectedErrorPaths.Count)
                $removedRows = @($rows | Where-Object Action -EQ Removed)
                (($removedRows.Path | Sort-Object) -join "`n") | Should-Be (($expectedRemoved | Sort-Object) -join "`n")
                @($removedRows | Where-Object Error) | Should-BeCollection -Count 0
                @($removedRows | Where-Object Type -EQ File) | Should-BeCollection -Count $e2eRemovedFiles.Count
                @($removedRows | Where-Object Type -EQ EmptyFolder) | Should-BeCollection -Count $e2eRemovedFolders.Count
                $errorRows = @($rows | Where-Object Error)
                (($errorRows.Path | Sort-Object) -join "`n") | Should-Be (($ExpectedErrorPaths | Sort-Object) -join "`n")
                @($errorRows | Where-Object Action) | Should-BeCollection -Count 0
            }
            else {
                @(Get-TestExcelFileHC) | Should-BeCollection -Count 0
            }
            @(Get-ChildItem -LiteralPath $testLogFolder -Filter '*System errors log*') | Should-BeCollection -Count 0
        }
    }

    It 'runs Example.json task <TaskIndex> with JobsTotal <JobsTotal> and verifies all files, folders and removal records' -ForEach @(
        foreach ($taskIndex in 0..8) {
            foreach ($jobsTotal in 1, 3) {
                @{ TaskIndex = $taskIndex; JobsTotal = $jobsTotal }
            }
        }
    ) {
        New-E2EFixtureHC
        $example = Get-Content -LiteralPath (Join-Path (Split-Path $PSScriptRoot) 'Example.json') -Raw | ConvertFrom-Json
        $example.Tasks | Should-BeCollection -Count 9
        $task = $example.Tasks[$TaskIndex]
        $task.ComputerName = 'localhost'
        $oldDate = (Get-Date).AddYears(-5)
        $futureDate = (Get-Date).AddYears(1)
        $null = Add-E2EFileHC -RelativePath 'outside-task\sentinel.txt' -CreationTime $oldDate -LastWriteTime $oldDate

        if ($task.Files) {
            $task.Files = @(
                for ($entryIndex = 0; $entryIndex -lt $task.Files.Count; $entryIndex++) {
                    $entry = $task.Files[$entryIndex]
                    $filePath = Add-E2EFileHC -RelativePath "files\selected-$entryIndex.txt" -CreationTime $oldDate -LastWriteTime $oldDate -Removed
                    if ($entry -is [string]) { $filePath }
                    else { [pscustomobject]@{ Name = $entry.Name; Path = $filePath } }
                }
            )
            $null = Add-E2EFileHC -RelativePath 'files\unselected.txt' -CreationTime $oldDate -LastWriteTime $oldDate
        }
        else {
            $task.Folders = @(
                for ($entryIndex = 0; $entryIndex -lt $task.Folders.Count; $entryIndex++) {
                    $entry = $task.Folders[$entryIndex]
                    $relativeRoot = "folder-$entryIndex"
                    $folderPath = Add-E2EFolderHC $relativeRoot
                    $hasAge = $null -ne $task.OlderThan
                    $removeRecent = $hasAge -and ($task.OlderThan.Quantity -eq 0)
                    $removeNested = $hasAge -and $task.Recurse
                    $null = Add-E2EFileHC "$relativeRoot\old.txt" -CreationTime $oldDate -LastWriteTime $oldDate -Removed:$hasAge
                    $null = Add-E2EFileHC "$relativeRoot\recent.txt" -CreationTime $futureDate -LastWriteTime $futureDate -Removed:$removeRecent
                    $null = Add-E2EFileHC "$relativeRoot\nested\old.txt" -CreationTime $oldDate -LastWriteTime $oldDate -Removed:$removeNested
                    $null = Add-E2EFolderHC "$relativeRoot\nested" -Removed:($removeNested -and $task.RemoveEmptyFolders)
                    $null = Add-E2EFolderHC "$relativeRoot\empty\deep" -Removed:$task.RemoveEmptyFolders
                    $null = Add-E2EFolderHC "$relativeRoot\empty" -Removed:$task.RemoveEmptyFolders

                    if ($task.Exclude.Folders) {
                        $task.Exclude.Folders = @(Add-E2EFolderHC "$relativeRoot\protected")
                        $null = Add-E2EFolderHC "$relativeRoot\protected\empty"
                        $null = Add-E2EFileHC "$relativeRoot\protected\history.json" -CreationTime $oldDate -LastWriteTime $oldDate
                    }
                    if ($task.Exclude.Files) {
                        $task.Exclude.Files = @(Add-E2EFileHC "$relativeRoot\State\last-run.json" -CreationTime $oldDate -LastWriteTime $oldDate)
                        $null = Add-E2EFileHC "$relativeRoot\State\other.json" -CreationTime $oldDate -LastWriteTime $oldDate -Removed
                    }
                    if ($entry -is [string]) { $folderPath }
                    else { [pscustomobject]@{ Name = $entry.Name; Path = $folderPath } }
                }
            )
        }

        $e2eConfiguration.Tasks = @($task)
        $e2eConfiguration.MaxConcurrent.JobsTotal = $JobsTotal
        Invoke-E2EAndAssertHC
        Should-NotInvoke New-PSSession -Scope It
        Should-NotInvoke Invoke-Command -Scope It
    }

    It 'uses the <Unit> calendar boundary and <BasedOn> for <ListName>' -ForEach @(
        foreach ($unit in 'Day', 'Month', 'Year') {
            foreach ($basedOn in 'CreationTime', 'LastWriteTime') {
                foreach ($listName in 'Files', 'Folders') {
                    @{ Unit = $unit; BasedOn = $basedOn; ListName = $listName }
                }
            }
        }
    ) {
        New-E2EFixtureHC
        Mock Get-Date { [datetime]'2026-10-02T12:00:00' }
        $boundary = switch ($Unit) {
            'Day' { [datetime]'2026-10-02T00:00:00' }
            'Month' { [datetime]'2026-10-01T00:00:00' }
            'Year' { [datetime]'2026-01-01T00:00:00' }
        }
        $selectedPaths = @(
            foreach ($case in @(
                    @{ Name = 'before'; Date = $boundary.AddSeconds(-1); Remove = $true }
                    @{ Name = 'boundary'; Date = $boundary; Remove = $false }
                    @{ Name = 'after'; Date = $boundary.AddSeconds(1); Remove = $false }
                )) {
                $dates = @{
                    CreationTime = if ($case.Remove) { [datetime]'2030-01-01' } else { [datetime]'2000-01-01' }
                    LastWriteTime = if ($case.Remove) { [datetime]'2030-01-01' } else { [datetime]'2000-01-01' }
                }
                $dates[$BasedOn] = $case.Date
                Add-E2EFileHC "selected\$($case.Name).txt" @dates -Removed:$case.Remove
            }
        )
        $null = Add-E2EFileHC 'outside\untouched.txt' -CreationTime ([datetime]'2000-01-01') -LastWriteTime ([datetime]'2000-01-01')
        $task = [pscustomobject]@{
            ComputerName = 'localhost'
            OlderThan = @{ Quantity = 1; Unit = $Unit; BasedOn = $BasedOn }
        }
        if ($ListName -eq 'Files') {
            $task | Add-Member Files $selectedPaths
        }
        else {
            $task | Add-Member Folders @(Join-Path $e2eRoot 'selected')
            $task | Add-Member Recurse $true
            $task | Add-Member RemoveEmptyFolders $true
        }
        $e2eConfiguration.Tasks = @($task)
        Invoke-E2EAndAssertHC
    }

    It 'honors normalized exclusions and exact names with quantity zero and JobsTotal <JobsTotal>' -ForEach @(1, 3 | ForEach-Object { @{ JobsTotal = $_ } }) {
        New-E2EFixtureHC
        $firstRoot = Add-E2EFolderHC 'first'
        $secondRoot = Add-E2EFolderHC 'second'
        $null = Add-E2EFileHC 'first\protected\history.json' -Hidden -ReadOnly
        $null = Add-E2EFolderHC 'first\protected\empty'
        $null = Add-E2EFileHC 'first\protected-other\remove.txt' -Removed
        $null = Add-E2EFolderHC 'first\protected-other' -Removed
        $null = Add-E2EFileHC 'first\state.json' -Hidden -ReadOnly
        $null = Add-E2EFileHC 'first\state.json.bak' -Removed
        $null = Add-E2EFileHC 'first\hidden\readonly.txt' -Hidden -ReadOnly -Removed
        $null = Add-E2EFolderHC 'first\hidden' -Removed
        $null = Add-E2EFileHC 'second\protected\remove.txt' -Removed
        $null = Add-E2EFolderHC 'second\protected' -Removed
        $null = Add-E2EFileHC 'second\state.json' -Removed
        $null = Add-E2EFolderHC 'second\empty' -Removed
        $e2eConfiguration.Tasks = @([pscustomobject]@{
                ComputerName = 'localhost'
                Folders = @("$firstRoot\unused\..", "$secondRoot\.")
                Exclude = @{
                    Folders = @("$firstRoot\unused\..\PROTECTED\")
                    Files = @("$firstRoot\STATE.JSON", "$firstRoot\state.json")
                }
                OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'LastWriteTime' }
                Recurse = $true
                RemoveEmptyFolders = $true
            })
        $e2eConfiguration.MaxConcurrent.JobsTotal = $JobsTotal
        Invoke-E2EAndAssertHC
    }

    It 'keeps nested files with Recurse false but removes empty folders at every depth with JobsTotal <JobsTotal>' -ForEach @(1, 3 | ForEach-Object { @{ JobsTotal = $_ } }) {
        New-E2EFixtureHC
        $rootPath = Add-E2EFolderHC 'selected'
        $null = Add-E2EFileHC 'selected\direct.txt' -Removed
        $null = Add-E2EFileHC 'selected\nested\old.txt' -CreationTime ((Get-Date).AddYears(-5))
        $null = Add-E2EFileHC 'selected\nested\hidden.txt' -Hidden
        $null = Add-E2EFolderHC 'selected\nested\empty\deep' -Removed
        $null = Add-E2EFolderHC 'selected\nested\empty' -Removed
        $null = Add-E2EFolderHC 'selected\empty' -Removed
        $e2eConfiguration.Tasks = @([pscustomobject]@{
                ComputerName = 'localhost'
                Folders = @($rootPath)
                OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'CreationTime' }
                Recurse = $false
                RemoveEmptyFolders = $true
            })
        $e2eConfiguration.MaxConcurrent.JobsTotal = $JobsTotal
        Invoke-E2EAndAssertHC
    }

    It 'finishes file tasks before empty-folder tasks even when listed in reverse order with JobsTotal <JobsTotal>' -ForEach @(1, 3 | ForEach-Object { @{ JobsTotal = $_ } }) {
        New-E2EFixtureHC
        $rootPath = Add-E2EFolderHC 'selected'
        $selectedFiles = @(
            Add-E2EFileHC 'selected\first\deep\remove.txt' -Removed
            Add-E2EFileHC 'selected\second\remove.txt' -Removed
        )
        $null = Add-E2EFolderHC 'selected\first\deep' -Removed
        $null = Add-E2EFolderHC 'selected\first' -Removed
        $null = Add-E2EFolderHC 'selected\second' -Removed
        $null = Add-E2EFileHC 'outside\sentinel.txt'
        $e2eConfiguration.Tasks = @(
            [pscustomobject]@{
                ComputerName = 'localhost'
                Folders = @($rootPath)
                RemoveEmptyFolders = $true
            }
            [pscustomobject]@{
                ComputerName = 'localhost'
                Files = $selectedFiles
                OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'LastWriteTime' }
            }
        )
        $e2eConfiguration.MaxConcurrent.JobsTotal = $JobsTotal
        Invoke-E2EAndAssertHC
    }

    It 'keeps all files and folders when nothing qualifies' {
        New-E2EFixtureHC
        $rootPath = Add-E2EFolderHC 'selected'
        $null = Add-E2EFileHC 'selected\recent.txt' -LastWriteTime ((Get-Date).AddYears(1))
        $null = Add-E2EFileHC 'selected\nested\recent.txt' -LastWriteTime ((Get-Date).AddYears(1))
        $null = Add-E2EFolderHC 'selected\empty'
        $e2eConfiguration.Tasks = @([pscustomobject]@{
                ComputerName = 'localhost'
                Folders = @($rootPath)
                OlderThan = @{ Quantity = 30; Unit = 'Day'; BasedOn = 'LastWriteTime' }
                Recurse = $true
                RemoveEmptyFolders = $false
            })
        Invoke-E2EAndAssertHC
    }

    It 'reports a missing <ListName> path and still processes valid sibling paths' -ForEach @('Files', 'Folders' | ForEach-Object { @{ ListName = $_ } }) {
        New-E2EFixtureHC
        $validFile = Add-E2EFileHC 'valid\remove.txt' -Removed
        $missingPath = Join-Path $e2eRoot 'missing'
        $task = [pscustomobject]@{
            ComputerName = 'localhost'
            OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'CreationTime' }
        }
        if ($ListName -eq 'Files') {
            $task | Add-Member Files @($missingPath, $validFile)
        }
        else {
            $task | Add-Member Folders @($missingPath, (Split-Path $validFile))
            $task | Add-Member Recurse $true
            $task | Add-Member RemoveEmptyFolders $false
        }
        $e2eConfiguration.Tasks = @($task)
        Invoke-E2EAndAssertHC -ExpectedErrorPaths @($missingPath)
    }

    It 'reports a locked file, keeps its parent and continues deleting other files and empty folders' {
        New-E2EFixtureHC
        $rootPath = Add-E2EFolderHC 'selected'
        $lockedPath = Add-E2EFileHC 'selected\locked\keep.txt'
        $null = Add-E2EFileHC 'selected\other\remove.txt' -Removed
        $null = Add-E2EFolderHC 'selected\other' -Removed
        $e2eConfiguration.Tasks = @([pscustomobject]@{
                ComputerName = 'localhost'
                Folders = @($rootPath)
                OlderThan = @{ Quantity = 0; Unit = 'Day'; BasedOn = 'LastWriteTime' }
                Recurse = $true
                RemoveEmptyFolders = $true
            })
        $lockStream = [System.IO.File]::Open($lockedPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        try {
            Invoke-E2EAndAssertHC -ExpectedErrorPaths @($lockedPath)
        }
        finally { $lockStream.Dispose() }
    }

    It 'passes remote job arguments through to the real worker using a local transport mock' {
        New-E2EFixtureHC
        $rootPath = Add-E2EFolderHC 'selected'
        $null = Add-E2EFileHC 'selected\old.txt' -LastWriteTime ((Get-Date).AddYears(-5)) -Removed
        $null = Add-E2EFileHC 'selected\recent.txt' -LastWriteTime ((Get-Date).AddYears(1))
        $null = Add-E2EFileHC 'selected\keep.json' -LastWriteTime ((Get-Date).AddYears(-5))
        $null = Add-E2EFolderHC 'selected\empty' -Removed
        Mock Invoke-Command {
            $FilePath | Should-Be (Join-Path (Split-Path $PSScriptRoot) 'Remove items.ps1')
            $ArgumentList[1] | Should-Be $rootPath
            & $FilePath @ArgumentList
        }
        $e2eConfiguration.Tasks = @([pscustomobject]@{
                ComputerName = 'E2E-MOCK-REMOTE'
                Folders = @($rootPath)
                Exclude = @{ Files = @(Join-Path $rootPath 'keep.json') }
                OlderThan = @{ Quantity = 30; Unit = 'Day'; BasedOn = 'LastWriteTime' }
                Recurse = $true
                RemoveEmptyFolders = $true
            })
        Invoke-E2EAndAssertHC
        Should-Invoke New-PSSession -Exactly -Times 2 -Scope It
        Should-Invoke Invoke-Command -Exactly -Times 2 -Scope It
        Should-Invoke Remove-PSSession -Exactly -Times 2 -Scope It
    }
}
Describe 'report inaccessible paths' {
    It 'counts repeated read errors once with file logging <SaveLogs>' -ForEach @(
        @{ SaveLogs = $true }
        @{ SaveLogs = $false }
    ) {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @($testNewInputFile.Tasks[1])
        $testNewInputFile.Tasks[0].RemoveEmptyFolders = $true
        $testNewInputFile.Settings.SendMail.When = 'OnError'
        if (-not $SaveLogs) { $testNewInputFile.Settings.SaveLogFiles.Where.Folder = $null }
        Mock Invoke-Command {
            [pscustomobject]@{
                DateTime = Get-Date
                ComputerName = 'PC2'
                Type = $ArgumentList[0]
                FullName = 'z:\folder\System Volume Information'
                CreationTime = $null
                Action = $null
                Error = 'Access denied'
            }
        }
        Test-NewJsonFileHC $testNewInputFile
        $global:LASTEXITCODE = 0

        $output = @(& $testScript @testParams -Verbose 4>&1)

        $LASTEXITCODE | Should-Be 1
        Should-Invoke Invoke-Command -Exactly -Times 2 -Scope It
        if ($SaveLogs) {
            $rows = @(Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName Overview)
            $rows | Should-BeCollection -Count 1
            $rows[0].Path | Should-Be 'z:\folder\System Volume Information'
            $rows[0].Error | Should-Be 'Access denied'
            $rows[0].Type | Should-Be 'FilesInFolder, EmptyFolders'
            $rows[0].OlderThan | Should-Be '1 Day'
            $rows[0].OlderThanBasedOn | Should-Be 'LastWriteTime'
        }
        else { @(Get-ChildItem -LiteralPath $testLogFolder) | Should-BeCollection -Count 0 }
        Should-Invoke Send-MailKitMessageHC -Exactly -Times 1 -Scope It -ParameterFilter {
            ($Subject -eq '0 removed, 1 error') -and
            ($Priority -eq 'High') -and
            ($Body -like '*>1 Error</span>*') -and
            ($Body -like '*0&nbsp;removed &middot; 1&nbsp;error*') -and
            ($Body -match "(?s)<tr class='path-row'>(?:(?!</tr>).)*class='removed-count'[^>]*>0</td>\s*<td class='error-count'[^>]*>1</td>")
        }
        Should-Invoke Write-EventLog -Exactly -Times 1 -Scope It -ParameterFilter {
            ($EntryType -eq 'Error') -and ($Message -like '*Access denied*')
        }
        @($output | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] }).Message |
        Should-ContainCollection 'Run summary: 0 removed, 1 errors'
    }
    It 'keeps the path error in Excel without a system errors log' {
        Clear-TestLogFolderHC
        Mock Invoke-Command {
            [pscustomobject]@{
                DateTime = Get-Date
                ComputerName = 'PC1'
                Type = 'File'
                FullName = 'z:\file.txt'
                CreationTime = $null
                Action = $null
                Error = 'Access to the path is denied.'
            }
        }
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @($testNewInputFile.Tasks[0])
        $testNewInputFile.Settings.SendMail.When = 'OnError'
        Test-NewJsonFileHC $testNewInputFile
        $global:LASTEXITCODE = 0

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 1
        $rows = @(Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName Overview)
        $rows | Should-BeCollection -Count 1
        $rows[0].Path | Should-Be 'z:\file.txt'
        $rows[0].Error | Should-Be 'Access to the path is denied.'
        $rows[0].Action | Should-BeNull
        @(Get-ChildItem -LiteralPath $testLogFolder -Filter '*System errors log.json') | Should-BeCollection -Count 0
        Should-Invoke Send-MailKitMessageHC -Exactly -Times 1 -Scope It -ParameterFilter {
            ($Priority -eq 'High') -and
            ($Subject -eq '0 removed, 1 error') -and
            ($Attachments -like '*Log.xlsx') -and
            (-not ($Attachments -like '*System errors log.json'))
        }
    }
    It 'preserves separate rows for <Case>' -ForEach @(
        @{ Case = 'different errors' }
        @{ Case = 'different paths' }
        @{ Case = 'separate input tasks' }
        @{ Case = 'successful removals' }
        @{ Case = 'deletion failures' }
    ) {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @($testNewInputFile.Tasks[1])
        if ($Case -eq 'separate input tasks') {
            $testNewInputFile.Tasks += Copy-ObjectHC $testNewInputFile.Tasks[0]
        }
        else { $testNewInputFile.Tasks[0].RemoveEmptyFolders = $true }

        Mock Invoke-Command {
            [pscustomobject]@{
                DateTime = Get-Date
                ComputerName = 'PC2'
                Type = if ($Case -eq 'deletion failures') { 'File' } else { $ArgumentList[0] }
                FullName = if (($Case -eq 'different paths') -and ($ArgumentList[0] -eq 'EmptyFolders')) {
                    'z:\folder\other'
                }
                else { 'z:\folder\denied' }
                CreationTime = $null
                Action = if ($Case -eq 'successful removals') { 'Removed' }
                Error = if ($Case -eq 'successful removals') { $null }
                elseif (($Case -eq 'different errors') -and ($ArgumentList[0] -eq 'EmptyFolders')) {
                    'Path no longer exists'
                }
                else { 'Access denied' }
            }
        }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $rows = @(Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName Overview)
        $rows | Should-BeCollection -Count 2
        $expectedSubject = if ($Case -eq 'successful removals') { '2 removed' } else { '0 removed, 2 errors' }
        Should-Invoke Send-MailKitMessageHC -Exactly -Times 1 -Scope It -ParameterFilter {
            $Subject -eq $expectedSubject
        }
    }
}
Describe 'create an Excel file' {
    BeforeAll {
        Clear-TestLogFolderHC

        Mock Invoke-Command {
            $testData[0]
        } -ParameterFilter {
            $ArgumentList[0] -eq 'File'
        }

        Mock Invoke-Command {
            $testData[1]
            $testData[2]
        } -ParameterFilter {
            $ArgumentList[0] -eq 'FilesInFolder'
        }

        Mock Invoke-Command {
            $testData[3]
        } -ParameterFilter {
            $ArgumentList[0] -eq 'EmptyFolders'
        }

        Test-NewJsonFileHC $testInputFile

        . $testScript @testParams

        $testExcelLogFile = Get-TestExcelFileHC
    }
    It 'in the log folder' {
        $testExcelLogFile | Should-NotBeNull
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
                    OlderThan    = "$($testInputFile.Tasks[0].OlderThan.Quantity) $($testInputFile.Tasks[0].OlderThan.Unit)"
                    Action       = $testData[0].Action
                    Error        = $testData[0].Error
                }
                @{
                    DateTime     = Get-Date
                    ComputerName = $testData[1].ComputerName
                    Type         = $testData[1].Type
                    Path         = $testData[1].FullName
                    CreationTime = $testData[1].CreationTime
                    OlderThan    = "$($testInputFile.Tasks[1].OlderThan.Quantity) $($testInputFile.Tasks[1].OlderThan.Unit)"
                    Action       = $testData[1].Action
                    Error        = $testData[1].Error
                }
                @{
                    DateTime     = Get-Date
                    ComputerName = $testData[2].ComputerName
                    Type         = $testData[2].Type
                    Path         = $testData[2].FullName
                    CreationTime = $testData[2].CreationTime
                    OlderThan    = "$($testInputFile.Tasks[1].OlderThan.Quantity) $($testInputFile.Tasks[1].OlderThan.Unit)"
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
            $actual | Should-BeCollection -Count $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Path -eq $testRow.Path
                }
                $actualRow.ComputerName | Should-Be $testRow.ComputerName
                $actualRow.Type | Should-Be $testRow.Type
                $actualRow.DateTime.ToString('yyyyMMdd') |
                Should-Be $testRow.DateTime.ToString('yyyyMMdd')
                $actualRow.CreationTime.ToString('yyyyMMdd HHmmss') |
                Should-Be $testRow.CreationTime.ToString('yyyyMMdd HHmmss')
                $actualRow.OlderThan | Should-Be $testRow.OlderThan
                $actualRow.Action | Should-Be $testRow.Action
                $actualRow.Error | Should-Be $testRow.Error
            }
        }
    }
    Context "with sheet 'Errors'" {
        BeforeAll {
            Clear-TestLogFolderHC

            Mock Invoke-Command {
                throw 'Oops'
            } -ParameterFilter {
                $ArgumentList[0] -eq 'File'
            }

            . $testScript @testParams

            $testExportedExcelRows = @(
                @{
                    ComputerName = $testInputFile.Tasks[0].ComputerName
                    Path         = $testInputFile.Tasks[0].Files[0].Path
                    Type         = 'RemoveFile'
                    OlderThan    = "$($testInputFile.Tasks[0].OlderThan.Quantity) $($testInputFile.Tasks[0].OlderThan.Unit)"
                    Error        = 'Oops'
                }
            )

            $actual = Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName 'Errors'
        }
        It 'with the correct total rows' {
            $actual | Should-BeCollection -Count $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            $testRow = $testExportedExcelRows[0]
            $actual.ComputerName | Should-Be $testRow.ComputerName
            $actual.Path | Should-Be $testRow.Path
            $actual.Type | Should-Be $testRow.Type
            $actual.OlderThan | Should-Be $testRow.OlderThan
            $actual.Error | Should-Be $testRow.Error
        }
        It 'includes execution diagnostics' {
            $actual.Stage | Should-Be 'Run remote worker'
            $actual.TargetObject | Should-Be 'Oops'
            $actual.FullyQualifiedErrorId | Should-Be 'Oops'
            $actual.ExceptionType | Should-Be 'System.Management.Automation.RuntimeException'
            $actual.ScriptStackTrace | Should-NotBeEmptyString
            $actual.PositionMessage | Should-BeLikeString '*throw*Oops*'
        }
        It 'does not duplicate execution or item errors in JSON' {
            @(Get-ChildItem -LiteralPath $testLogFolder -Filter '*System errors log.json') | Should-BeCollection -Count 0
        }
    }
}
Describe 'job failure diagnostics' {
    BeforeEach {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @($testNewInputFile.Tasks[0])
        $testNewInputFile.Settings.SendMail.When = 'Never'
        $testNewInputFile.Settings.SaveInEventLog.Save = $false
        $global:LASTEXITCODE = 0
    }
    It 'distinguishes session setup from worker execution' {
        Mock New-PSSession { throw 'Session setup failed' }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 1
        $rows = @(Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName Errors)
        $rows | Should-BeCollection -Count 1
        $rows[0].Stage | Should-Be 'Open remote session'
        $rows[0].Error | Should-Be 'Session setup failed'
        $rows[0].FullyQualifiedErrorId | Should-Be 'Session setup failed'
        Should-NotInvoke Invoke-Command -Scope It
        @(Get-ChildItem -LiteralPath $testLogFolder -Filter '*System errors log.json') | Should-BeCollection -Count 0
    }
    It 'distinguishes the failed target from the configured path' {
        Mock Invoke-Command {
            Write-Error -Exception ([System.IO.FileNotFoundException]::new('The system cannot find the file specified.')) -ErrorId 'MissingChild' -Category ObjectNotFound -TargetObject 'z:\child.txt' -ErrorAction Stop
        }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 1
        $rows = @(Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName Errors)
        $rows | Should-BeCollection -Count 1
        $rows[0].Path | Should-Be 'z:\file.txt'
        $rows[0].TargetObject | Should-Be 'z:\child.txt'
        $rows[0].Stage | Should-Be 'Run remote worker'
        $rows[0].ExceptionType | Should-Be 'System.IO.FileNotFoundException'
        $rows[0].FullyQualifiedErrorId | Should-BeLikeString '*MissingChild*'
        $rows[0].Error | Should-Be 'The system cannot find the file specified.'
        $rows[0].ScriptStackTrace | Should-NotBeEmptyString
        @(Get-ChildItem -LiteralPath $testLogFolder -Filter '*System errors log.json') | Should-BeCollection -Count 0
    }
    It 'preserves real worker diagnostics with JobsTotal <_>' -ForEach @(1, 3) {
        $testFile = New-Item "TestDrive:/diagnostics-$_.txt" -ItemType File
        $testNewInputFile.MaxConcurrent.JobsTotal = $_
        $testNewInputFile.Tasks[0].ComputerName = 'localhost'
        $testNewInputFile.Tasks[0].Files = @($testFile.FullName)
        $testNewInputFile.Tasks[0].OlderThan.Unit = 'Year'
        $testNewInputFile.Tasks[0].OlderThan.Quantity = [int]::MaxValue
        Test-NewJsonFileHC $testNewInputFile

        & $testScript -ConfigurationJsonFile $testOutParams.FilePath

        $LASTEXITCODE | Should-Be 1
        (Test-Path -LiteralPath $testFile.FullName) | Should-BeTrue
        $rows = @(Import-Excel -Path (Get-TestExcelFileHC).FullName -WorksheetName Errors)
        $rows | Should-BeCollection -Count 1
        $rows[0].Stage | Should-Be 'Run local worker'
        $rows[0].Error | Should-BeLikeString '*Invalid retention period*'
        $rows[0].ExceptionType | Should-Be 'System.Management.Automation.RuntimeException'
        $rows[0].FullyQualifiedErrorId | Should-BeLikeString '*Invalid retention period*'
        $rows[0].ScriptStackTrace | Should-BeLikeString '*Get-ExclusiveCutoffHC*Remove items.ps1*'
        $rows[0].PositionMessage | Should-BeLikeString '*Remove items.ps1*'
        @(Get-ChildItem -LiteralPath $testLogFolder -Filter '*System errors log.json') | Should-BeCollection -Count 0
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

            Should-NotInvoke Send-MailKitMessageHC -Scope It
        }
    }
    Context 'send an e-mail when' {
        It "'OnError' and there are errors" {
            Mock Invoke-Command { $testData[1] }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Settings.SendMail.When = 'OnError'

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It
        }
        It "'OnErrorOrAction' and there are actions but no errors" {
            Mock Invoke-Command { $testData[0] }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Settings.SendMail.When = 'OnErrorOrAction'

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It
        }
        It "'OnErrorOrAction' and there are errors but no actions" {
            Mock Invoke-Command { $testData[1] }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Settings.SendMail.When = 'OnErrorOrAction'

            Test-NewJsonFileHC $testNewInputFile

            .$testScript @testParams

            Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope It
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
        $testNewInputFile.Tasks = @($testNewInputFile.Tasks[0])

        Test-NewJsonFileHC $testNewInputFile

        . $testScript @testParams
    }
    It 'with the Settings.SendMail properties' {
        Should-Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
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
        Should-Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($Priority -eq 'High') -and
            ($Subject -eq '1 removed, 1 error') -and
            ($Attachments -like '*Log.xlsx') -and
            (-not ($Attachments -like '* - System errors log.json'))
        }
    }
    It 'saves the exact email body with the same log prefix without attaching it' {
        $testExcelFile = Get-TestExcelFileHC
        $testHtmlPath = Join-Path $testLogFolder ($testExcelFile.Name.Replace(' - Log.xlsx', ' - Mail.html'))
        (Test-Path -LiteralPath $testHtmlPath) | Should-BeTrue
        $testHtmlBody = Get-Content -LiteralPath $testHtmlPath -Raw -Encoding utf8
        $testBrowserUrl = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $testHtmlPath))
        $testHtmlBody | Should-BeLikeString '*If this mail is not visible*'
        $testHtmlBody.Contains("href='$testBrowserUrl'") | Should-BeTrue

        Should-Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($Body -ceq $testHtmlBody) -and
            (-not ($Attachments -like '* - Mail.html'))
        }
    }
    It 'with the correct body' {
        Should-Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($Body -like '*<h1>Test (Brecht)</h1>*') -and
            ($Body -like '*Email body*') -and
            # a card for the computer, with a link to the file over its admin share
            ($Body -like '*>PC1</p>*') -and
            ($Body -like "*href='file:////PC1/z$/file.txt'*>FTP log file</a>*") -and
            ($Body -like '*Remove file older than 1 day (creation time)*') -and
            ($Body -match "(?s)<tr class='path-row'>(?:(?!</tr>).)*class='removed-count'[^>]*>1</td>\s*<td class='error-count'[^>]*>1</td>") -and
            ($Body -like '*Started*Ended*Duration*')
        }
    }
    It 'with Settings.SendMail.Subject added to the subject' {
        $testNewInputFile.Settings.SendMail.Subject = 'Custom'
        Test-NewJsonFileHC $testNewInputFile

        . $testScript @testParams

        Should-Invoke Send-MailKitMessageHC -Exactly 1 -Scope It -ParameterFilter {
            $Subject -eq '1 removed, 1 error, Custom'
        }
    }
}
Describe 'email rows per input path' {
    It 'summarizes exclusions once across paths without repeating the cleanup rule' {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @(
            [pscustomobject]@{
                ComputerName = 'PC1'
                Folders = @('z:\first', 'z:\second')
                OlderThan = @{ Quantity = 1; Unit = 'Day'; BasedOn = 'LastWriteTime' }
                Recurse = $true
                RemoveEmptyFolders = $false
                Exclude = @{ Folders = @('z:\first\keep'); Files = @('z:\second\keep.txt') }
            }
        )
        Mock Invoke-Command
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $testHtmlFile = @(Get-ChildItem -LiteralPath $testLogFolder -Filter '* - Mail.html')
        $testHtmlFile | Should-BeCollection -Count 1
        $html = Get-Content -LiteralPath $testHtmlFile[0].FullName -Raw
        [regex]::Matches($html, 'Remove files older than 1 day').Count | Should-Be 1
        $html | Should-BeLikeString '*including subfolders, excluding 1 folder and 1 file</caption>*'
        [regex]::Matches($html, "class='path-row'").Count | Should-Be 2
    }
    It 'separates paths and combines cleanup phases without merging separate tasks (<Label>)' -ForEach @(
        @{ Label = 'local paths'; FirstPath = 'z:\first'; SecondPath = 'z:\second'; ComputerLabel = 'PC1' }
        @{ Label = 'multiple UNC servers'; FirstPath = '\\SERVER1\Logs'; SecondPath = '\\SERVER2\Logs'; ComputerLabel = 'PC1' }
    ) {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @(
            [pscustomobject]@{
                ComputerName = 'PC1'
                Folders = @(
                    @{ Name = 'First & primary'; Path = $FirstPath }
                    @{ Name = 'Second'; Path = $SecondPath }
                )
                OlderThan = @{ Quantity = 1; Unit = 'Day'; BasedOn = 'LastWriteTime' }
                Recurse = $true
                RemoveEmptyFolders = $true
            }
            [pscustomobject]@{
                ComputerName = 'PC1'
                Folders = @(@{ Name = 'Separate task'; Path = $FirstPath })
                RemoveEmptyFolders = $true
            }
        )
        Mock Invoke-Command {
            if (($ArgumentList[0] -eq 'EmptyFolders') -and ($ArgumentList[1] -eq $SecondPath)) {
                throw 'Folder job failed'
            }
            [pscustomobject]@{
                DateTime = Get-Date
                ComputerName = 'PC1'
                Type = $ArgumentList[0]
                FullName = $ArgumentList[1]
                Action = if ($ArgumentList[0] -eq 'FilesInFolder') { 'Removed' }
                Error = if ($ArgumentList[0] -eq 'EmptyFolders') { 'Access denied' }
            }
        }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        Should-Invoke Invoke-Command -Exactly -Times 5 -Scope It
        $testHtmlFile = @(Get-ChildItem -LiteralPath $testLogFolder -Filter '* - Mail.html')
        $testHtmlFile | Should-BeCollection -Count 1
        $html = Get-Content -LiteralPath $testHtmlFile[0].FullName -Raw
        [regex]::Matches($html, 'class="task-table"').Count | Should-Be 2
        [regex]::Matches($html, '>First &amp; primary</a>').Count | Should-Be 1
        [regex]::Matches($html, '>Second</a>').Count | Should-Be 1
        [regex]::Matches($html, '>Separate task</a>').Count | Should-Be 1
        $expectedRoots = if ($Label -eq 'local paths') { 1 } else { 0 }
        [regex]::Matches($html, "class='root-breadcrumb'").Count | Should-Be $expectedRoots
        $html | Should-BeLikeString '*Remove files older than 1 day (last write time), including subfolders; Remove empty folders*'
        [regex]::Matches($html, 'Remove files older than 1 day').Count | Should-Be 1
        $rows = [regex]::Matches($html, "(?s)<tr class='path-row'>.*?</tr>").Value
        $rows | Should-BeCollection -Count 3
        ($rows | Where-Object { $_ -like '*>First &amp; primary</a>*' }) | Should-BeLikeString "*class='removed-count'*>1</td>*class='error-count'*>1</td>*"
        ($rows | Where-Object { $_ -like '*>Second</a>*' }) | Should-BeLikeString "*class='removed-count'*>1</td>*class='error-count'*>1</td>*"
        ($rows | Where-Object { $_ -like '*>Separate task</a>*' }) | Should-BeLikeString "*class='removed-count'*>0</td>*class='error-count'*>1</td>*"
        $html | Should-BeLikeString "*>$ComputerLabel</p>*"
        $html | Should-BeLikeString '*2 paths</p>*'
        Should-Invoke Send-MailKitMessageHC -Exactly -Times 1 -Scope It -ParameterFilter {
            ($Subject -eq '2 removed, 3 errors') -and ($Body -ceq $html)
        }
    }
}
Describe 'email rows for locally executed UNC files' {
    It 'shows individual file counters under the local host in both email and saved HTML' {
        Clear-TestLogFolderHC
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @(
            [pscustomobject]@{
                ComputerName = $null
                Files = @(
                    @{ Name = 'Removed file'; Path = '\\SERVER1\Logs\removed.txt' }
                    @{ Name = 'Failed file'; Path = '\\SERVER2\Logs\failed.txt' }
                    @{ Name = 'Unchanged file'; Path = '\\SERVER2\Logs\unchanged.txt' }
                )
                OlderThan = @{ Quantity = 1; Unit = 'Day'; BasedOn = 'CreationTime' }
            }
        )
        $testWorker = Join-Path $TestDrive 'mail-file-worker.ps1'
        Set-Content -LiteralPath $testWorker -Value @'
$path = $args[1]
if ($path -notlike '*\unchanged.txt') {
    [pscustomobject]@{
        DateTime = Get-Date
        ComputerName = $env:COMPUTERNAME
        Type = 'File'
        FullName = $path
        Action = if ($path -like '*\removed.txt') { 'Removed' }
        Error = if ($path -like '*\failed.txt') { 'Access denied' }
    }
}
'@
        Test-NewJsonFileHC $testNewInputFile

        & $testScript -ConfigurationJsonFile $testOutParams.FilePath -RemoveItemsScript $testWorker

        Should-NotInvoke New-PSSession -Scope It
        Should-NotInvoke Invoke-Command -Scope It
        $testHtmlFile = @(Get-ChildItem -LiteralPath $testLogFolder -Filter '* - Mail.html')
        $testHtmlFile | Should-BeCollection -Count 1
        $html = Get-Content -LiteralPath $testHtmlFile[0].FullName -Raw
        [regex]::Matches($html, 'class="task-table"').Count | Should-Be 1
        [regex]::Matches($html, 'Remove file older than 1 day').Count | Should-Be 1
        [regex]::Matches($html, "class='root-breadcrumb'").Count | Should-Be 1
        $html | Should-BeLikeString '*\\SERVER2\Logs\</td>*'
        $html | Should-NotBeLikeString '*\\SERVER1\Logs\</td>*'
        $rows = [regex]::Matches($html, "(?s)<tr class='path-row'>.*?</tr>").Value
        $rows | Should-BeCollection -Count 3
        ($rows | Where-Object { $_ -like '*>Removed file</a>*' }) | Should-BeLikeString "*class='removed-count'*>1</td>*class='error-count'*>0</td>*"
        ($rows | Where-Object { $_ -like '*>Failed file</a>*' }) | Should-BeLikeString "*class='removed-count'*>0</td>*class='error-count'*>1</td>*"
        ($rows | Where-Object { $_ -like '*>Unchanged file</a>*' }) | Should-BeLikeString "*class='removed-count'*>0</td>*class='error-count'*>0</td>*"
        $html | Should-BeLikeString "*>$env:COMPUTERNAME</p>*"
        $html | Should-NotBeLikeString '*>SERVER1</p>*'
        $html | Should-NotBeLikeString '*>SERVER2</p>*'
        $html | Should-BeLikeString '*3 paths</p>*'
        $html | Should-BeLikeString "*href='file:////SERVER1/Logs/removed.txt'*"
        Should-Invoke Send-MailKitMessageHC -Exactly -Times 1 -Scope It -ParameterFilter {
            ($Subject -eq '1 removed, 1 error') -and ($Body -ceq $html)
        }
    }
}
Describe 'save email HTML' {
    BeforeEach {
        Clear-TestLogFolderHC
        Mock Invoke-Command
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @($testNewInputFile.Tasks[0])
        $testNewInputFile.Settings.SaveInEventLog.Save = $false
        $global:LASTEXITCODE = 0
    }
    It 'does not save HTML when <_> suppresses email' -ForEach @('Never', 'OnError', 'OnErrorOrAction') {
        $testNewInputFile.Settings.SendMail.When = $_
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 0
        @(Get-ChildItem -LiteralPath $testLogFolder -Filter '* - Mail.html') | Should-BeCollection -Count 0
        Should-NotInvoke Send-MailKitMessageHC -Scope It
    }
    It 'does not write HTML when file logging is disabled' {
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = $null
        Mock Set-Content -ParameterFilter { $LiteralPath -like '* - Mail.html' }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 0
        Should-NotInvoke Set-Content -Scope It -ParameterFilter { $LiteralPath -like '* - Mail.html' }
        Should-Invoke Send-MailKitMessageHC -Exactly -Times 1 -Scope It -ParameterFilter {
            $Body -notlike '*If this mail is not visible*'
        }
    }
    It 'retains the HTML when sending fails' {
        Mock Send-MailKitMessageHC { throw 'SMTP unavailable' }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 1
        @(Get-ChildItem -LiteralPath $testLogFolder -Filter '* - Mail.html') | Should-BeCollection -Count 1
        ((Get-TestSystemErrorsHC).Message -join "`n") | Should-BeLikeString '*Failed sending email: SMTP unavailable*'
    }
    It 'still sends email and logs the error when saving HTML fails' {
        Mock Set-Content { throw 'HTML write denied' } -ParameterFilter { $LiteralPath -like '* - Mail.html' }
        Test-NewJsonFileHC $testNewInputFile

        & $testScript @testParams

        $LASTEXITCODE | Should-Be 1
        Should-Invoke Send-MailKitMessageHC -Exactly -Times 1 -Scope It -ParameterFilter {
            $Body -notlike '*If this mail is not visible*'
        }
        ((Get-TestSystemErrorsHC).Message -join "`n") | Should-BeLikeString '*Failed saving email HTML: HTML write denied*'
    }
}
Describe 'Settings.SaveInEventLog' {
    BeforeAll {
        Mock Invoke-Command { throw 'Oops' }
    }
    It 'writes job errors to the event log when Save is true' {
        Test-NewJsonFileHC $testInputFile

        .$testScript @testParams

        Should-Invoke Write-EventLog -Scope It -ParameterFilter {
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

        Should-NotInvoke Write-EventLog -Scope It
    }
}
Describe 'Settings.SaveLogFiles.DeleteLogsAfterDays' {
    It 'removes log files older than the given days' {
        $testOldLogFile = New-Item -Path "$testLogFolder\old log.txt" -ItemType File -Force
        $testOldLogFile.LastWriteTime = (Get-Date).AddDays(-3)

        $testNewLogFile = New-Item -Path "$testLogFolder\new log.txt" -ItemType File -Force

        Test-NewJsonFileHC $testInputFile

        .$testScript @testParams

        (Test-Path -LiteralPath $testOldLogFile.FullName) | Should-BeFalse
        (Test-Path -LiteralPath $testNewLogFile.FullName) | Should-BeTrue
    }
}
