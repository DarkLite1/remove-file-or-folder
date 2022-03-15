#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
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
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
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
            It 'Remove is missing' {
                @{
                    MailTo       = @('bob@contoso.com')
                    Destinations = @(
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
                    MailTo       = @('bob@contoso.com')
                    Destinations = @(
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
                        MailTo       = @('bob@contoso.com')
                        Destinations = @(
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
                                Remove             = 'file'
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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*'OlderThanDays' needs to be a number, the value 'a' is not supported*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'RemoveEmptyFolders is not null' {
                    @{
                        MailTo       = @('bob@contoso.com')
                        Destinations = @(
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
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile* 'RemoveEmptyFolders' cannot be used with 'Remove' value 'file'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context "Remove is 'folder'" {
                It 'RemoveEmptyFolders is not null' {
                    @{
                        MailTo       = @('bob@contoso.com')
                        Destinations = @(
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
    BeforeAll {
        $testFolder = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
        }
        $testFile = 0..2 | ForEach-Object {
            (New-Item "TestDrive:/file$_.txt" -ItemType File).FullName
        }

        @{
            MailTo       = @('bob@contoso.com')
            Destinations = @(
                @{
                    Remove        = 'file'
                    Path          = $testFile[0]
                    ComputerName  = $env:COMPUTERNAME
                    OlderThanDays = 0
                }
                @{
                    Remove        = 'file'
                    Path          = 'c:\notExistingFileOrFolder'
                    ComputerName  = $env:COMPUTERNAME
                    OlderThanDays = 0
                }
            )
        } | ConvertTo-Json | Out-File @testOutParams

        $testExportedExcelRows = @(
            @{
                ComputerName = $env:COMPUTERNAME
                Type         = 'File'
                Path         = $testFile[0]
                Error        = $null
                Action       = 'Removed'
            }
            @{
                ComputerName = $env:COMPUTERNAME
                Type         = 'File'
                Path         = 'c:\notExistingFileOrFolder'
                Error        = 'Path not found'
                Action       = $null
            }
        )
        
        $Error.Clear()
        . $testScript @testParams
    }
    Context 'remove the requested' {
        It 'file' {
            $testFile[0] | Should -Not -Exist
        }
    }
    Context 'not remove other' {
        It 'folders' {
            $testFolder[0] | Should -Exist
            $testFolder[1] | Should -Exist
            $testFolder[2] | Should -Exist
        }
        It 'files' {
            $testFile[1] | Should -Exist
            $testFile[2] | Should -Exist
        }
    } 
    Context 'exports an Excel file' {
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
            foreach ($testResult in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Path -eq $testResult.Path
                }
                $actualRow.ComputerName | Should -Be $testResult.ComputerName
                $actualRow.Type | Should -Be $testResult.Type
                $actualRow.Path | Should -Be $testResult.Path
                $actualRow.Error | Should -Be $testResult.Error
                $actualRow.Action | Should -Be $testResult.Action
            }
        }
    }
    It 'sends a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq 'High') -and
            # ($Subject -eq 'Removed 4/5 items, 1 removal errors') -and
            ($Attachments -like '*log.xlsx') -and
            ($Message -like '*
            *Successfully removed items*1*
            *Errors while removing items*1*
            *Not existing items after running the script*2*')
        }
    } -Tag test
} 