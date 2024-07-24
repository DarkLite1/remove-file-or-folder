#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testInputFile = @{
        SendMail          = @{
            To   = 'bob@contoso.com'
            When = 'Always'
        }
        MaxConcurrentJobs = 1
        Remove            = @{
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
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        Path        = @{
            RemoveFileScript          = (New-Item 'TestDrive:/b.ps1' -ItemType File).FullName
            RemoveEmptyFoldersScript  = (New-Item 'TestDrive:/a.ps1' -ItemType File).FullName
            RemoveFilesInFolderScript = (New-Item 'TestDrive:/c.ps1' -ItemType File).FullName
        }
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Invoke-Command
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
    Context 'the file is not found' {
        It 'Path.<_>' -ForEach @(
            'RemoveEmptyFoldersScript', 'RemoveFile', 'RemoveFilesInFolder'
        ) {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.Path.$_ = 'c:\NotExisting.ps1'

            $testInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*Path.$_ 'c:\NotExisting.ps1' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
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
            It '<_> not found' -ForEach @(
                'SendMail', 'MaxConcurrentJobs', 'Remove'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.When not supported' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.When = 'NotSupported'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*Value '$($testNewInputFile.SendMail.When)' for 'SendMail.When' is not supported. Supported values are 'Never, OnlyOnError, OnlyOnErrorOrAction or Always'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context "Remove.File" {
                BeforeEach {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Remove = @{
                        File = $testNewInputFile.Remove.File
                    }
                }
                It '<_> not found' -ForEach @(
                    'Path', 'OlderThan'
                ) {
                    $testNewInputFile.Remove.File[0].$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Remove.File.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                Context 'OlderThan' {
                    Context 'OlderThan.Unit' {
                        It 'not found' {
                            $testNewInputFile.Remove.File[0].OlderThan.Remove("Unit")

                            $testNewInputFile | ConvertTo-Json -Depth 5 |
                            Out-File @testOutParams

                            .$testScript @testParams

                            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Remove.File.OlderThan.Unit' found*")
                            }
                            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                                $EntryType -eq 'Error'
                            }
                        }
                        It 'is not supported' {
                            $testNewInputFile.Remove.File[0].OlderThan.Unit = 'notSupported'

                            $testNewInputFile | ConvertTo-Json -Depth 5 |
                            Out-File @testOutParams

                            .$testScript @testParams

                            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'notSupported' is not supported by 'Remove.File.OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'*")
                            }
                            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                                $EntryType -eq 'Error'
                            }
                        }
                    }
                    Context 'OlderThan.Quantity' {
                        It 'not found' {
                            $testNewInputFile.Remove.File[0].OlderThan.Remove("Quantity")

                            $testNewInputFile | ConvertTo-Json -Depth 5 |
                            Out-File @testOutParams

                            .$testScript @testParams

                            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Remove.File.OlderThan.Quantity' not found. Use value number '0' to move all files*")
                            }
                            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                                $EntryType -eq 'Error'
                            }
                        }
                        It 'is not a number' {
                            $testNewInputFile.Remove.File[0].OlderThan.Quantity = 'a'

                            $testNewInputFile | ConvertTo-Json -Depth 5 |
                            Out-File @testOutParams

                            .$testScript @testParams

                            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Remove.File.OlderThan.Quantity' needs to be a number, the value 'a' is not supported*")
                            }
                            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                                $EntryType -eq 'Error'
                            }
                        }
                    }
                }
                It 'Path is a local path but no ComputerName is given' {
                    $testNewInputFile.Remove.File[0].ComputerName = $null
                    $testNewInputFile.Remove.File[0].Path = 'd:\bla'

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*No 'Remove.File.ComputerName' found for path '$($testNewInputFile.Remove.File[0].Path)'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context "Remove.FilesInFolder" {
                BeforeEach {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Remove = @{
                        FilesInFolder = $testNewInputFile.Remove.FilesInFolder
                    }
                }
                It '<_> not found' -ForEach @(
                    'Path', 'OlderThan'
                ) {
                    $testNewInputFile.Remove.FilesInFolder[0].$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Remove.FilesInFolder.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                Context 'OlderThan' {
                    Context 'OlderThan.Unit' {
                        It 'not found' {
                            $testNewInputFile.Remove.FilesInFolder[0].OlderThan.Remove("Unit")

                            $testNewInputFile | ConvertTo-Json -Depth 5 |
                            Out-File @testOutParams

                            .$testScript @testParams

                            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Remove.FilesInFolder.OlderThan.Unit' found*")
                            }
                            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                                $EntryType -eq 'Error'
                            }
                        }
                        It 'is not supported' {
                            $testNewInputFile.Remove.FilesInFolder[0].OlderThan.Unit = 'notSupported'

                            $testNewInputFile | ConvertTo-Json -Depth 5 |
                            Out-File @testOutParams

                            .$testScript @testParams

                            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'notSupported' is not supported by 'Remove.FilesInFolder.OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'*")
                            }
                            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                                $EntryType -eq 'Error'
                            }
                        }
                    }
                    Context 'OlderThan.Quantity' {
                        It 'not found' {
                            $testNewInputFile.Remove.FilesInFolder[0].OlderThan.Remove("Quantity")

                            $testNewInputFile | ConvertTo-Json -Depth 5 |
                            Out-File @testOutParams

                            .$testScript @testParams

                            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Remove.FilesInFolder.OlderThan.Quantity' not found. Use value number '0' to move all files*")
                            }
                            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                                $EntryType -eq 'Error'
                            }
                        }
                        It 'is not a number' {
                            $testNewInputFile.Remove.FilesInFolder[0].OlderThan.Quantity = 'a'

                            $testNewInputFile | ConvertTo-Json -Depth 5 |
                            Out-File @testOutParams

                            .$testScript @testParams

                            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Remove.FilesInFolder.OlderThan.Quantity' needs to be a number, the value 'a' is not supported*")
                            }
                            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                                $EntryType -eq 'Error'
                            }
                        }
                    }
                }
                It 'Recurse it not a boolean' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Remove.FilesInFolder[0].Recurse = 'a'

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*Property 'Remove.FilesInFolder.Recurse' is not a boolean value*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Path is a local path but no ComputerName is given' {
                    $testNewInputFile.Remove.FilesInFolder[0].ComputerName = $null
                    $testNewInputFile.Remove.FilesInFolder[0].Path = 'd:\bla'

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*No 'Remove.FilesInFolder.ComputerName' found for path '$($testNewInputFile.Remove.FilesInFolder[0].Path)'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context "Remove.EmptyFolders" {
                BeforeEach {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Remove = @{
                        EmptyFolders = $testNewInputFile.Remove.EmptyFolders
                    }
                }
                It '<_> not found' -ForEach @(
                    'Path'
                ) {
                    $testNewInputFile.Remove.EmptyFolders[0].$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Remove.EmptyFolders.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Path is a local path but no ComputerName is given' {
                    $testNewInputFile.Remove.EmptyFolders[0].ComputerName = $null
                    $testNewInputFile.Remove.EmptyFolders[0].Path = 'd:\bla'

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*No 'Remove.EmptyFolders.ComputerName' found for path '$($testNewInputFile.Remove.EmptyFolders[0].Path)'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe 'execute script' {
    Context 'Remove.File' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile

            $testNewInputFile.Remove = @{
                File = $testNewInputFile.Remove.File
            }

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq $testNewInputFile.Remove.File[0].ComputerName) -and
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

            $testNewInputFile.Remove = @{
                FilesInFolder = $testNewInputFile.Remove.FilesInFolder
            }

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq $testNewInputFile.Remove.FilesInFolder[0].ComputerName) -and
                ($FilePath -eq $testParams.Path.RemoveFilesInFolderScript) -and
                ($ArgumentList[0] -eq $testNewInputFile.Remove.FilesInFolder[0].Path) -and
                ($ArgumentList[1] -eq $testNewInputFile.Remove.FilesInFolder[0].OlderThan.Unit) -and
                ($ArgumentList[2] -eq $testNewInputFile.Remove.FilesInFolder[0].OlderThan.Quantity) -and
                ($ArgumentList[3] -eq $testNewInputFile.Remove.FilesInFolder[0].Recurse)
            }
        }
    }
    Context 'Remove.RemoveEmptyFolders' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile

            $testNewInputFile.Remove = @{
                EmptyFolders = $testNewInputFile.Remove.EmptyFolders
            }

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'with the correct arguments' {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq $testNewInputFile.Remove.EmptyFolders[0].ComputerName) -and
                ($FilePath -eq $testParams.Path.RemoveEmptyFoldersScript) -and
                ($ArgumentList[0] -eq $testNewInputFile.Remove.EmptyFolders[0].Path)
            }
        }
    }
}