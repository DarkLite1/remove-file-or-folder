#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $commandImportExcel = Get-Command Import-Excel

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        MailTo     = @('bob@contoso.com')
        Path       = New-Item 'TestDrive:/folders.xlsx' -ItemType File
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('Path', 'MailTo', 'ScriptName') {
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
    It 'the path to the Excel file is not found' {
        $testNewParams = $testParams.clone()
        $testNewParams.Path = 'notFound.xlsx'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*'notFound.xlsx'*file not found*")
        }
    }
    Context 'The column header is not found in the Excel sheet' {
        It 'ComputerName' {
            Mock Import-Excel {
                @(
                    [PSCustomObject]@{
                        NoComputerName = 'PC1'
                        Path           = 'K:\folder1'
                    }
                )
            }
            Mock Invoke-Command

            . $testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Column headers 'ComputerName' and 'Path' are not found*")
            }
        }
        It 'Path' {
            Mock Import-Excel {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        FullName     = 'K:\folder1'
                    }
                )
            }
            Mock Invoke-Command

            . $testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Column headers 'ComputerName' and 'Path' are not found*")
            }
        }
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
        It 'a summary mail is sent to the user' {
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
}