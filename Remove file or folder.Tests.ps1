#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testInputFile = @{
        MailTo       = @('bob@contoso.com')
        Destinations = @(
            @{
                Remove             = 'content'
                Path               = '\\contoso\share'
                ComputerName       = $null
                OlderThanDays      = 20
                RemoveEmptyFolders = $false
            }
        )
    }
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }
    $testInputFile | ConvertTo-Json | Out-File @testOutParams

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
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
            AfterAll {
                $testInputFile | ConvertTo-Json | Out-File @testOutParams
            }
            It 'Destinations is missing' {
                @{
                    MailTo = @('bob@contoso.com')
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
                    MailTo       = @('bob@contoso.com')
                    Destinations = @(
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
                    MailTo       = @('bob@contoso.com')
                    Destinations = @(
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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'ComputerName' found for path 'd:\bla'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context "Remove is 'content'" {
                It 'RemoveEmptyFolders is missing' {
                    @{
                        MailTo       = @('bob@contoso.com')
                        Destinations = @(
                            @{
                                Remove             = 'content'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                OlderThanDays      = 20
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'RemoveEmptyFolders' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'RemoveEmptyFolders is not a boolean' {
                    @{
                        MailTo       = @('bob@contoso.com')
                        Destinations = @(
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
                        MailTo       = @('bob@contoso.com')
                        Destinations = @(
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
                        MailTo       = @('bob@contoso.com')
                        Destinations = @(
                            @{
                                Remove             = 'content'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                RemoveEmptyFolders = $false
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'OlderThanDays' number found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'OlderThanDays is not a number' {
                    @{
                        MailTo       = @('bob@contoso.com')
                        Destinations = @(
                            @{
                                Remove             = 'content'
                                Path               = '\\contoso\share'
                                ComputerName       = $null
                                OlderThanDays      = 'a'
                                RemoveEmptyFolders = $false
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*'OlderThanDays' needs to be a number, the value 'a' is not supported*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context "Remove is 'folder'" {

                
            }
        } #-Tag test
    }
}
Describe 'when rows are imported from Excel' {
    Context 'Invoke-Command' {
        BeforeAll {
            Mock Import-Excel {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        Path         = 'K:\folder1'
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        Path         = 'K:\folder2'
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        Path         = 'K:\folder3'
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC3'
                        Path         = $null
                    }
                )
            }
            Mock Invoke-Command

            . $testScript @testParams
        }
        It 'is only called once for each computer' {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq 'PC1') -and
                ($ArgumentList[0][0] -eq 'K:\folder1') -and
                ($ArgumentList[0][1] -eq 'K:\folder2')
            }
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ComputerName -eq 'PC2') -and
                ($ArgumentList[0] -eq 'K:\folder3')
            }
        }
        It 'is not called when there are no paths for a computer' {
            Should -Not -Invoke Invoke-Command -Scope Context -ParameterFilter {
                ($ComputerName -eq 'PC3') 
            }
        }
    } 
    Context 'the script' {
        BeforeAll {
            $testFolder = 0..2 | ForEach-Object {
                (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
            }
            $testFile = 0..2 | ForEach-Object {
                (New-Item "TestDrive:/file$_" -ItemType File).FullName
            }

            Mock Import-Excel {
                @(
                    [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        Path         = $testFolder[0]
                    }
                    [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        Path         = $testFolder[1]
                    }
                    [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        Path         = $testFile[0]
                    }
                    [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        Path         = $testFile[1]
                    }
                    [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        Path         = 'notExistingFileOrFolder'
                    }
                )
            }
    
            $Error.Clear()
            . $testScript @testParams
        }
        Context 'starts a job on the remote computer to' {
            Describe 'remove the requested' {
                It 'folders' {
                    $testFolder[0] | Should -Not -Exist
                    $testFolder[1] | Should -Not -Exist
                }
                It 'files' {
                    $testFile[0] | Should -Not -Exist
                    $testFile[1] | Should -Not -Exist
                }
            }
            Describe 'not remove other' {
                It 'folders' {
                    $testFolder[2] | Should -Exist
                }
                It 'files' {
                    $testFile[2] | Should -Exist
                }
            }
        }
        Context 'exports an Excel file' {
            BeforeAll {
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                $actual = & $commandImportExcel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the same quantity of rows as the imported Excel worksheet' {
                $actual | Should -HaveCount 5
            }
            It 'with the successful removals' {
                @{
                    0 = $testFolder[0]
                    1 = $testFolder[1]
                    2 = $testFile[0]
                    3 = $testFile[1]
                }.GetEnumerator() | ForEach-Object {
                    $actual[$_.Key].ComputerName | Should -Be $env:COMPUTERNAME
                    $actual[$_.Key].Path | Should -Be $_.Value
                    $actual[$_.Key].Date | Should -Not -BeNullOrEmpty
                    $actual[$_.Key].Exist | Should -BeFalse
                    $actual[$_.Key].Action | Should -Be 'Remove'
                    $actual[$_.Key].Error | Should -BeNullOrEmpty    
                }
            }
            It 'with the failed removals' {
                $actual[4].ComputerName | Should -Be $env:COMPUTERNAME
                $actual[4].Path | Should -Be 'notExistingFileOrFolder'
                $actual[4].Date | Should -Not -BeNullOrEmpty
                $actual[4].Exist | Should -BeFalse
                $actual[4].Action | Should -BeNullOrEmpty
                $actual[4].Error | Should -Be 'Path not found'
            }
        }
        It 'sends a summary mail to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq 'Removed 4/5 items, 1 removal errors') -and
                ($Attachments -like '*log.xlsx') -and
                ($Message -like '*
                *Successfully removed items*4*
                *Errors while removing items*1*
                *Imported Excel file rows*5*
                *Not existing items after running the script*5*')
            }
        }
    }
} -Skip