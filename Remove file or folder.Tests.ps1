#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $realCmdLet = @{
        StartJob      = Get-Command Start-Job
        InvokeCommand = Get-Command Invoke-Command
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It 'MailTo is missing' {
                @{
                    # MailTo       = @('bob@contoso.com')
                    MaxConcurrentJobs = 4
                    Destinations      = @()
                } | ConvertTo-Json | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'MailTo' addresses found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Destinations is missing' {
                @{
                    MailTo            = @('bob@contoso.com')
                    MaxConcurrentJobs = 4
                } | ConvertTo-Json | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Destinations' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Path is missing' {
                @{
                    MailTo            = @('bob@contoso.com')
                    MaxConcurrentJobs = 4
                    Destinations      = @(
                        @{
                            ComputerName       = $null
                            OlderThanDays      = 'a'
                            RemoveEmptyFolders = $false
                        }
                    )
                } | ConvertTo-Json | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Path' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Path is a local path but no ComputerName is given' {
                @{
                    MailTo            = @('bob@contoso.com')
                    MaxConcurrentJobs = 4
                    Destinations      = @(
                        @{
                            Path               = 'd:\bla'
                            ComputerName       = $null
                            OlderThanDays      = 'a'
                            RemoveEmptyFolders = $false
                        }
                    )
                } | ConvertTo-Json | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*destination path 'd:\bla'*No 'ComputerName' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Remove is missing' {
                @{
                    MailTo            = @('bob@contoso.com')
                    MaxConcurrentJobs = 4
                    Destinations      = @(
                        @{
                            Path               = '\\contoso\share'
                            ComputerName       = $null
                            OlderThanDays      = 0
                            RemoveEmptyFolders = $false
                        }
                    )
                } | ConvertTo-Json | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Remove' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Remove value is incorrect' {
                @{
                    MailTo            = @('bob@contoso.com')
                    MaxConcurrentJobs = 4
                    Destinations      = @(
                        @{
                            Remove             = 'wrong'
                            Path               = '\\contoso\share'
                            ComputerName       = $null
                            OlderThanDays      = 0
                            RemoveEmptyFolders = $false
                        }
                    )
                } | ConvertTo-Json | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'wrong' in 'Remove' is not valid, only values 'folder', 'file' or 'content' are supported*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context "Remove is 'content'" {
                It 'RemoveEmptyFolders is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 4
                        Destinations      = @(
                            @{
                                Remove        = 'content'
                                Path          = '\\contoso\share'
                                ComputerName  = $null
                                OlderThanDays = 20
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'RemoveEmptyFolders' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'RemoveEmptyFolders is not a boolean' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 4
                        Destinations      = @(
                            @{
                                Remove             = 'content'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                OlderThanDays      = 20
                                RemoveEmptyFolders = 'yes'
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*The value 'yes' in 'RemoveEmptyFolders' is not a true false value*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'RemoveEmptyFolders is correct' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 4
                        Destinations      = @(
                            @{
                                Remove             = 'content'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                OlderThanDays      = 20
                                RemoveEmptyFolders = $false
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams

                    Should -Not -Invoke Send-MailHC -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*'RemoveEmptyFolders'*")
                    }
                }
            }
            Context "Remove is 'file'" {
                It 'OlderThanDays is missing' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 4
                        Destinations      = @(
                            @{
                                Remove             = 'file'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                RemoveEmptyFolders = $false
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'OlderThanDays' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'OlderThanDays is not a number' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 4
                        Destinations      = @(
                            @{
                                Remove             = 'file'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                OlderThanDays      = 'a'
                                RemoveEmptyFolders = $false
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'OlderThanDays' needs to be a number, the value 'a' is not supported*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'RemoveEmptyFolders is not null' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 4
                        Destinations      = @(
                            @{
                                Remove             = 'file'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                OlderThanDays      = 0
                                RemoveEmptyFolders = $true
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile* Property 'RemoveEmptyFolders' cannot be used with 'Remove' value 'file'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context "Remove is 'folder'" {
                It 'RemoveEmptyFolders is not null' {
                    @{
                        MailTo            = @('bob@contoso.com')
                        MaxConcurrentJobs = 4
                        Destinations      = @(
                            @{
                                Remove             = 'folder'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                OlderThanDays      = 0
                                RemoveEmptyFolders = $true
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile* 'RemoveEmptyFolders' cannot be used with 'Remove' value 'folder'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe "when 'Remove' is 'file'" {
    Context  "and 'OlderThanDays' is '0'" {
        BeforeAll {
            $testFolder = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
            }
            $testFile = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/file$_.txt" -ItemType File).FullName
            }

            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 4
                Destinations      = @(
                    @{
                        Name          = 'FTP log file'
                        Remove        = 'file'
                        Path          = $testFile[0]
                        ComputerName  = $env:COMPUTERNAME
                        OlderThanDays = 0
                    }
                    @{
                        Remove        = 'file'
                        Path          = 'c:\Not Existing File'
                        ComputerName  = $env:COMPUTERNAME
                        OlderThanDays = 0
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            $testExportedExcelRows = @(
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Type          = 'File'
                    Path          = $testFile[0]
                    Error         = $null
                    Action        = 'Removed'
                    OlderThanDays = 0
                }
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Type          = 'File'
                    Path          = 'c:\not existing file'
                    Error         = 'Path not found'
                    Action        = $null
                    OlderThanDays = 0
                }
            )
            $testRemoved = @{
                files   = @($testFile[0])
                folders = $null
            }
            $testNotRemoved = @{
                files   = @($testFile[1], $testFile[2])
                folders = @($testFolder[0], $testFolder[1], $testFolder[2])
            }
            $testMail = @{
                Priority = 'High'
                Subject  = '1 removed, 1 error'
                Message  = "*<ul><li><a href=`"\\$env:COMPUTERNAME\c$\not existing file`">\\$env:COMPUTERNAME\c$\not existing file</a><br>Remove file<br>Removed: 0, <b style=`"color:red;`">errors: 1</b><br><br></li>*<li><a href=`"*$($testFile[0].Name)`">FTP log file</a><br>Remove file<br>Removed: 1</li></ul><p><i>* Check the attachment for details</i></p>*"
            }

            $Error.Clear()
            . $testScript @testParams
        }
        Context 'remove the requested' {
            It 'files' {
                $testRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
            It 'folders' {
                $testRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
        }
        Context 'not remove other' {
            It 'files' {
                $testNotRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
            It 'folders' {
                $testNotRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
        }
        Context 'export an Excel file' {
            BeforeAll {
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
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
                    $actualRow.OlderThanDays | Should -Be $testRow.OlderThanDays
                    $actualRow.Path | Should -Be $testRow.Path
                    $actualRow.Error | Should -Be $testRow.Error
                    $actualRow.Action | Should -Be $testRow.Action
                }
            }
        }
        It 'send a summary mail to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq $testMail.Priority) -and
            ($Subject -eq $testMail.Subject) -and
            ($Attachments -like '*log.xlsx') -and
            ($Message -like $testMail.Message)
            }
        }
    }
    Context  "and 'OlderThanDays' is not '0'" {
        BeforeAll {
            $testFolder = 0..2 | ForEach-Object {
                (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
            }
            $testFile = 0..2 | ForEach-Object {
                (New-Item "TestDrive:/file$_.txt" -ItemType File).FullName
            }

            @($testFile[0], $testFolder[0]) | ForEach-Object {
                $testItem = Get-Item -LiteralPath $_
                $testItem.CreationTime = (Get-Date).AddDays(-5)
            }

            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 4
                Destinations      = @(
                    @{
                        Remove        = 'file'
                        Path          = $testFile[0]
                        ComputerName  = $env:COMPUTERNAME
                        OlderThanDays = 3
                    }
                    @{
                        Remove        = 'file'
                        Path          = $testFile[1]
                        ComputerName  = $env:COMPUTERNAME
                        OlderThanDays = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            $testRemoved = @{
                files   = @($testFile[0])
                folders = $null
            }
            $testNotRemoved = @{
                files   = @($testFile[1], $testFile[2])
                folders = @($testFolder[0], $testFolder[1], $testFolder[2])
            }

            . $testScript @testParams
        }
        Context 'remove the requested' {
            It 'files' {
                $testRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
            It 'folders' {
                $testRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
        }
        Context 'not remove other' {
            It 'files' {
                $testNotRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
            It 'folders' {
                $testNotRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
        }
    }
}
Describe "when 'Remove' is 'folder'" {
    Context  "and 'OlderThanDays' is '0'" {
        BeforeAll {
            $testFolder = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
            }
            $testFile = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/file$_.txt" -ItemType File).FullName
            }

            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 4
                Destinations      = @(
                    @{
                        Remove        = 'folder'
                        Path          = $testFolder[0]
                        ComputerName  = $env:COMPUTERNAME
                        OlderThanDays = 0
                    }
                    @{
                        Remove        = 'folder'
                        Path          = 'c:\Not Existing Folder'
                        ComputerName  = $env:COMPUTERNAME
                        OlderThanDays = 0
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            $testExportedExcelRows = @(
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Type          = 'Folder'
                    Path          = $testFolder[0]
                    Error         = $null
                    Action        = 'Removed'
                    OlderThanDays = 0
                }
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Type          = 'Folder'
                    Path          = 'c:\not existing folder'
                    Error         = 'Path not found'
                    Action        = $null
                    OlderThanDays = 0
                }
            )
            $testRemoved = @{
                files   = $null
                folders = @($testFolder[0])
            }
            $testNotRemoved = @{
                files   = @($testFile[0], $testFile[1], $testFile[2])
                folders = @($testFolder[1], $testFolder[2])
            }
            $testMail = @{
                Priority = 'High'
                Subject  = '1 removed, 1 error'
                Message  = "*<ul><li><a href=`"\\$env:COMPUTERNAME\c$\not existing folder`">\\$env:COMPUTERNAME\c$\not existing folder</a><br>Remove folder<br>Removed: 0, <b style=`"color:red;`">errors: 1</b><br><br></li>*$($testFolder[0].Name)*Remove folder<br>Removed: 1</li></ul><p><i>* Check the attachment for details</i></p>*"
            }

            $Error.Clear()
            . $testScript @testParams
        }
        Context 'remove the requested' {
            It 'files' {
                $testRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
            It 'folders' {
                $testRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
        }
        Context 'not remove other' {
            It 'files' {
                $testNotRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
            It 'folders' {
                $testNotRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
        }
        Context 'export an Excel file' {
            BeforeAll {
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
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
                    $actualRow.Path | Should -Be $testRow.Path
                    $actualRow.OlderThanDays | Should -Be $testRow.OlderThanDays
                    $actualRow.Error | Should -Be $testRow.Error
                    $actualRow.Action | Should -Be $testRow.Action
                }
            }
        }
        It 'send a summary mail to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq $testMail.Priority) -and
            ($Subject -eq $testMail.Subject) -and
            ($Attachments -like '*log.xlsx') -and
            ($Message -like $testMail.Message)
            }
        }
    }
    Context  "and 'OlderThanDays' is not '0'" {
        BeforeAll {
            $testFolder = 0..2 | ForEach-Object {
                (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
            }
            $testFile = 0..2 | ForEach-Object {
                (New-Item "TestDrive:/file$_.txt" -ItemType File).FullName
            }

            @($testFile[0], $testFolder[0]) | ForEach-Object {
                $testItem = Get-Item -LiteralPath $_
                $testItem.CreationTime = (Get-Date).AddDays(-5)
            }

            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 4
                Destinations      = @(
                    @{
                        Remove        = 'folder'
                        Path          = $testFolder[0]
                        ComputerName  = $env:COMPUTERNAME
                        OlderThanDays = 3
                    }
                    @{
                        Remove        = 'folder'
                        Path          = $testFolder[1]
                        ComputerName  = $env:COMPUTERNAME
                        OlderThanDays = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            $testRemoved = @{
                files   = $null
                folders = @($testFolder[0])
            }
            $testNotRemoved = @{
                files   = @($testFile[0], $testFile[1], $testFile[2])
                folders = @($testFolder[1], $testFolder[2])
            }

            . $testScript @testParams
        }
        Context 'remove the requested' {
            It 'files' {
                $testRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
            It 'folders' {
                $testRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
        }
        Context 'not remove other' {
            It 'files' {
                $testNotRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
            It 'folders' {
                $testNotRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
        }
    }
}
Describe "when 'Remove' is 'content' and remove empty folders" {
    Context  "and 'OlderThanDays' is '0'" {
        BeforeAll {
            $testFolder = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
            }
            $testFile = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/file$_.txt" -ItemType File).FullName
            }

            $testFolder +=
        (New-Item "$($testFolder[0])/sub" -ItemType Directory).FullName

            $testFile +=
        (New-Item "$($testFolder[0])/sub/file.txt" -ItemType File).FullName

            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 4
                Destinations      = @(
                    @{
                        Remove             = 'content'
                        Path               = $testFolder[0]
                        ComputerName       = $env:COMPUTERNAME
                        RemoveEmptyFolders = $true
                        OlderThanDays      = 0
                    }
                    @{
                        Remove             = 'content'
                        Path               = 'c:\Not Existing Folder'
                        ComputerName       = $env:COMPUTERNAME
                        RemoveEmptyFolders = $true
                        OlderThanDays      = 0
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            $testExportedExcelRows = @(
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Type          = 'Folder'
                    Path          = $testFolder[3]
                    Error         = $null
                    Action        = 'Removed'
                    OlderThanDays = 0
                }
                @{
                    ComputerName  = $env:COMPUTERNAME
                    Type          = 'File'
                    Path          = $testFile[3]
                    Error         = $null
                    Action        = 'Removed'
                    OlderThanDays = 0
                }
            )
            $testRemoved = @{
                files   = @($testFile[3])
                folders = @($testFolder[3])
            }
            $testNotRemoved = @{
                files   = @($testFile[0], $testFile[1], $testFile[2])
                folders = @($testFolder[0], $testFolder[1], $testFolder[2])
            }

            $Error.Clear()
            . $testScript @testParams
        }
        Context 'remove the requested' {
            It 'files' {
                $testRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
            It 'folders' {
                $testRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
        }
        Context 'not remove other' {
            It 'files' {
                $testNotRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
            It 'folders' {
                $testNotRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
        }
        Context 'export an Excel file' {
            BeforeAll {
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
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
                    $actualRow.Path | Should -Be $testRow.Path
                    $actualRow.OlderThanDays | Should -Be $testRow.OlderThanDays
                    $actualRow.Error | Should -Be $testRow.Error
                    $actualRow.Action | Should -Be $testRow.Action
                    $actualRow.CreationTime | Should -Not -BeNullOrEmpty
                }
            }
        }
        Context 'Send a mail to the user' {
            BeforeAll {
                $testMail = @{
                    To          = 'bob@contoso.com'
                    Bcc         = $ScriptAdmin
                    Priority    = 'High'
                    Subject     = '2 removed, 1 error'
                    Message     = "*<ul><li><a href=`"\\$env:COMPUTERNAME\c$\not existing folder`">\\$env:COMPUTERNAME\c$\not existing folder</a><br>Remove folder content and remove empty folders<br>Removed: 0, <b style=`"color:red;`">errors: 1</b><br><br></li>*$($testFolder[0].Name)*Remove folder content and remove empty folders<br>Removed: 2</li></ul><p><i>* Check the attachment for details</i></p>*"
                    Attachments = '* - log.xlsx'
                }
            }
            It 'Send-MailHC has the correct arguments' {
                $mailParams.To | Should -Be $testMail.To
                $mailParams.Bcc | Should -Be $testMail.Bcc
                $mailParams.Subject | Should -Be $testMail.Subject
                $mailParams.Message | Should -BeLike $testMail.Message
                $mailParams.Attachments | Should -BeLike $testMail.Attachments
            }
            It 'Send-MailHC is called' {
                Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq $testMail.To) -and
                ($Bcc -eq $testMail.Bcc) -and
                ($Priority -eq $testMail.Priority) -and
                ($Subject -eq $testMail.Subject) -and
                ($Attachments -like $testMail.Attachments) -and
                ($Message -like $testMail.Message)
                }
            }
        }
    }
    Context  "and 'OlderThanDays' is not '0'" {
        BeforeAll {
            $testFolder = @(
                'TestDrive:/folderA' ,
                'TestDrive:/folderB' ,
                'TestDrive:/folderA/subA',
                'TestDrive:/folderA/subAA',
                'TestDrive:/folderB/subB',
                'TestDrive:/folderB/subBB'
            ) | ForEach-Object {
                (New-Item $_ -ItemType Directory).FullName
            }
            $testFile = @(
                'TestDrive:/fileX.txt',
                'TestDrive:/fileZ.txt'
                'TestDrive:/folderA/fileA.txt',
                'TestDrive:/folderA/subA/fileSubA.txt',
                'TestDrive:/folderB/fileB.txt' ,
                'TestDrive:/folderB/subB/fileSubB.txt'
            ) | ForEach-Object {
                (New-Item $_ -ItemType File).FullName
            }

            @(
                $testFolder[0],
                $testFolder[2],
                $testFile[0],
                $testFile[1],
                $testFile[2],
                $testFile[4],
                $testFile[5]
            ) | ForEach-Object {
                $testItem = Get-Item -LiteralPath $_
                $testItem.CreationTime = (Get-Date).AddDays(-5)
            }

            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 4
                Destinations      = @(
                    @{
                        Remove             = 'content'
                        Path               = $testFolder[0]
                        ComputerName       = $env:COMPUTERNAME
                        OlderThanDays      = 3
                        RemoveEmptyFolders = $true
                    }
                    @{
                        Remove             = 'content'
                        Path               = $testFolder[1]
                        ComputerName       = $env:COMPUTERNAME
                        OlderThanDays      = 3
                        RemoveEmptyFolders = $true
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            $testRemoved = @{
                files   = @($testFile[2], $testFile[4], $testFile[5])
                folders = @($testFolder[3], $testFolder[4], $testFolder[5])
            }
            $testNotRemoved = @{
                files   = @($testFile[3], $testFile[0], $testFile[1])
                folders = @($testFolder[0], $testFolder[2])
            }

            . $testScript @testParams
        }
        Context 'remove the requested' {
            It 'files' {
                $testRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
            It 'folders' {
                $testRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
        }
        Context 'not remove other' {
            It 'files' {
                $testNotRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
            It 'folders' {
                $testNotRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
        }
    }
}
Describe "when 'Remove' is 'content' and do not remove empty folders" {
    Context  "and 'OlderThanDays' is '0'" {
        BeforeAll {
            $testFolder = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
            }
            $testFile = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/file$_.txt" -ItemType File).FullName
            }

            $testFile +=
        (New-Item "$($testFolder[0])/file.txt" -ItemType File).FullName

            $testFolder +=
        (New-Item "$($testFolder[0])/sub" -ItemType Directory).FullName

            $testFile +=
        (New-Item "$($testFolder[0])/sub/file.txt" -ItemType File).FullName

            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 4
                Destinations      = @(
                    @{
                        Remove             = 'content'
                        Path               = $testFolder[0]
                        ComputerName       = $env:COMPUTERNAME
                        RemoveEmptyFolders = $false
                        OlderThanDays      = 0
                    }
                    @{
                        Remove             = 'content'
                        Path               = 'c:\Not Existing Folder'
                        ComputerName       = $env:COMPUTERNAME
                        RemoveEmptyFolders = $true
                        OlderThanDays      = 0
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            $testRemoved = @{
                files   = @($testFile[3], $testFile[4])
                folders = $null
            }
            $testNotRemoved = @{
                files   = @($testFile[0], $testFile[1], $testFile[2])
                folders = @(
                    $testFolder[0], $testFolder[1], $testFolder[2], $testFolder[3]
                )
            }

            . $testScript @testParams
        }
        Context 'remove the requested' {
            It 'files' {
                $testRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
            It 'folders' {
                $testRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
        }
        Context 'not remove other' {
            It 'files' {
                $testNotRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
            It 'folders' {
                $testNotRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
        }
        Context 'export an Excel file' {
            Context 'with worksheet Overview' {
                BeforeAll {
                    $testExportedExcelRows = @(
                        @{
                            ComputerName  = $env:COMPUTERNAME
                            Type          = 'File'
                            Path          = $testFile[3]
                            Error         = $null
                            Action        = 'Removed'
                            OlderThanDays = 0
                        }
                        @{
                            ComputerName  = $env:COMPUTERNAME
                            Type          = 'File'
                            Path          = $testFile[4]
                            Error         = $null
                            Action        = 'Removed'
                            OlderThanDays = 0
                        }
                    )

                    $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                    $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
                }
                It 'to the log folder' {
                    $testExcelLogFile | Should -Not -BeNullOrEmpty
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
                        $actualRow.OlderThanDays | Should -Be $testRow.OlderThanDays
                        $actualRow.Path | Should -Be $testRow.Path
                        $actualRow.Error | Should -Be $testRow.Error
                        $actualRow.Action | Should -Be $testRow.Action
                        $actualRow.CreationTime | Should -Not -BeNullOrEmpty
                    }
                }
            }
            Context 'with worksheet Errors' {
                BeforeAll {
                    $testExportedExcelRows = @(
                        @{
                            ComputerName       = $env:COMPUTERNAME
                            Path               = 'c:\Not Existing Folder'
                            Remove             = 'content'
                            RemoveEmptyFolders = $true
                            OlderThanDays      = 0
                            Error              = "Folder 'c:\not existing folder' not found"
                        }
                    )

                    $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                    $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Errors'
                }
                It 'to the log folder' {
                    $testExcelLogFile | Should -Not -BeNullOrEmpty
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
                        $actualRow.Path | Should -Be $testRow.Path
                        $actualRow.Remove | Should -Be $testRow.Remove
                        $actualRow.OlderThanDays | Should -Be $testRow.OlderThanDays
                        $actualRow.RemoveEmptyFolders |
                        Should -Be $testRow.RemoveEmptyFolders
                        $actualRow.Error | Should -Be $testRow.Error
                    }
                }
            }
        }
        Context 'Send a mail to the user' {
            BeforeAll {
                $testMail = @{
                    To          = 'bob@contoso.com'
                    Bcc         = $ScriptAdmin
                    Priority    = 'High'
                    Subject     = '2 removed, 1 error'
                    Message     = "*<ul><li><a href=`"\\$env:COMPUTERNAME\c$\not existing folder`">\\$env:COMPUTERNAME\c$\not existing folder</a><br>Remove folder content and remove empty folders<br>Removed: 0, <b style=`"color:red;`">errors: 1</b><br><br></li>*$($testFolder[0].Name)*Remove folder content<br>Removed: 2</li></ul><p><i>* Check the attachment for details</i></p>*"
                    Attachments = '* - log.xlsx'
                }
            }
            It 'Send-MailHC has the correct arguments' {
                $mailParams.To | Should -Be $testMail.To
                $mailParams.Bcc | Should -Be $testMail.Bcc
                $mailParams.Subject | Should -Be $testMail.Subject
                $mailParams.Message | Should -BeLike $testMail.Message
                $mailParams.Attachments | Should -BeLike $testMail.Attachments
            }
            It 'Send-MailHC is called' {
                Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq $testMail.To) -and
                ($Bcc -eq $testMail.Bcc) -and
                ($Priority -eq $testMail.Priority) -and
                ($Subject -eq $testMail.Subject) -and
                ($Attachments -like $testMail.Attachments) -and
                ($Message -like $testMail.Message)
                }
            }
        }
    }
    Context  "and 'OlderThanDays' is not '0'" {
        BeforeAll {
            $testFolder = @(
                'TestDrive:/folderA' ,
                'TestDrive:/folderB' ,
                'TestDrive:/folderA/subA',
                'TestDrive:/folderB/subB'
            ) | ForEach-Object {
                (New-Item $_ -ItemType Directory).FullName
            }
            $testFile = @(
                'TestDrive:/fileX.txt',
                'TestDrive:/fileZ.txt'
                'TestDrive:/folderA/fileA.txt',
                'TestDrive:/folderA/subA/fileSubA.txt',
                'TestDrive:/folderB/fileB.txt' ,
                'TestDrive:/folderB/subB/fileSubB.txt'
            ) | ForEach-Object {
                (New-Item $_ -ItemType File).FullName
            }

            @(
                $testFolder[0],
                $testFolder[2],
                $testFile[0],
                $testFile[1],
                $testFile[2],
                $testFile[4],
                $testFile[5]
            ) | ForEach-Object {
                $testItem = Get-Item -LiteralPath $_
                $testItem.CreationTime = (Get-Date).AddDays(-5)
            }

            @{
                MailTo            = @('bob@contoso.com')
                MaxConcurrentJobs = 4
                Destinations      = @(
                    @{
                        Remove             = 'content'
                        Path               = $testFolder[0]
                        ComputerName       = $env:COMPUTERNAME
                        OlderThanDays      = 3
                        RemoveEmptyFolders = $false
                    }
                    @{
                        Remove             = 'content'
                        Path               = $testFolder[1]
                        ComputerName       = $env:COMPUTERNAME
                        OlderThanDays      = 3
                        RemoveEmptyFolders = $false
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            $testRemoved = @{
                files   = @($testFile[2], $testFile[4], $testFile[5])
                folders = $null
            }
            $testNotRemoved = @{
                files   = @($testFile[3], $testFile[0], $testFile[1])
                folders = @($testFolder[0], $testFolder[2])
            }

            . $testScript @testParams
        }
        Context 'remove the requested' {
            It 'files' {
                $testRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
            It 'folders' {
                $testRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Not -Exist
                }
            }
        }
        Context 'not remove other' {
            It 'files' {
                $testNotRemoved.files | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
            It 'folders' {
                $testNotRemoved.folders | Where-Object { $_ } | ForEach-Object {
                    $_ | Should -Exist
                }
            }
        }
    }
}
Describe 'a non terminating job error' {
    BeforeAll {
        $testFolder = (New-Item 'TestDrive:/folder' -ItemType Directory).FullName

        @{
            MailTo            = @('bob@contoso.com')
            MaxConcurrentJobs = 4
            Destinations      = @(
                @{
                    Remove             = 'content'
                    Path               = $testFolder
                    ComputerName       = $env:COMPUTERNAME
                    RemoveEmptyFolders = $false
                    OlderThanDays      = 0
                }
            )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        $testExportedExcelRows = @(
            @{
                ComputerName       = $env:COMPUTERNAME
                Path               = $testFolder
                Remove             = 'content'
                RemoveEmptyFolders = $false
                OlderThanDays      = 0
                Error              = 'Oops'
            }
        )

        Mock Invoke-Command {
            & $realCmdLet.InvokeCommand -Scriptblock {
                Write-Error 'Oops'
            } -AsJob -ComputerName $env:COMPUTERNAME
        }

        . $testScript @testParams
    }
    Context 'export to Excel in worksheet Errors' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Errors'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
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
                $actualRow.Path | Should -Be $testRow.Path
                $actualRow.Remove | Should -Be $testRow.Remove
                $actualRow.OlderThanDays | Should -Be $testRow.OlderThanDays
                $actualRow.RemoveEmptyFolders |
                Should -Be $testRow.RemoveEmptyFolders
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
    Context 'Send a mail to the user' {
        BeforeAll {
            $testMail = @{
                To          = 'bob@contoso.com'
                Bcc         = $ScriptAdmin
                Priority    = 'High'
                Subject     = '0 removed, 1 error'
                Message     = "*<ul><li><a href=`"\\$env:COMPUTERNAME\c$\*\folder`">\\$env:COMPUTERNAME\c$\*\folder</a><br>Remove folder content<br>Removed: 0, <b style=`"color:red;`">errors: 1</b></li></ul><p><i>* Check the attachment for details</i></p>*"
                Attachments = '* - log.xlsx'
            }
        }
        It 'Send-MailHC has the correct arguments' {
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeLike $testMail.Attachments
        }
        It 'Send-MailHC is called' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testMail.To) -and
            ($Bcc -eq $testMail.Bcc) -and
            ($Priority -eq $testMail.Priority) -and
            ($Subject -eq $testMail.Subject) -and
            ($Attachments -like $testMail.Attachments) -and
            ($Message -like $testMail.Message)
            }
        }
    }
}