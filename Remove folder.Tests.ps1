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
}
Describe 'when rows are imported from Excel' {
    Context 'Invoke-Command' {
        BeforeAll {
            Mock Import-Excel {
                @(
                    [PSCustomObject]@{
                        PSComputerName = 'PC1'
                        FullName       = 'K:\folder1'
                    }
                    [PSCustomObject]@{
                        PSComputerName = 'PC1'
                        FullName       = 'K:\folder2'
                    }
                    [PSCustomObject]@{
                        PSComputerName = 'PC2'
                        FullName       = 'K:\folder3'
                    }
                    [PSCustomObject]@{
                        PSComputerName = 'PC3'
                        FullName       = $null
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
    Context 'folders on the computer are' {
        BeforeAll {
            $testFolder = 0..2 | ForEach-Object {
                (New-Item "TestDrive:/folder$_" -ItemType Directory).FullName
            }
            $testFile = (New-Item 'TestDrive:/file1' -ItemType File).FullName

            Mock Import-Excel {
                @(
                    [PSCustomObject]@{
                        PSComputerName = $env:COMPUTERNAME
                        FullName       = $testFolder[0]
                    }
                    [PSCustomObject]@{
                        PSComputerName = $env:COMPUTERNAME
                        FullName       = $testFolder[1]
                    }
                    [PSCustomObject]@{
                        PSComputerName = $env:COMPUTERNAME
                        FullName       = 'notExistingFolder'
                    }
                    [PSCustomObject]@{
                        PSComputerName = $env:COMPUTERNAME
                        FullName       = $testFile
                    }
                )
            }
    
            . $testScript @testParams
        }
        It 'removed when requested' {
            $testFolder[0] | Should -Not -Exist
            $testFolder[1] | Should -Not -Exist
        } 
        It 'not removed when not requested' {
            $testFolder[2] | Should -Exist
        }
        It 'a file name instead of a folder does not remove the file' {
            $testFile | Should -Exist
        }
        Context 'the exported Excel file' {
            BeforeAll {
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

                $actual = & $commandImportExcel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
            }
            It 'is saved in the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'contains the same quantity of rows as the request' {
                $actual | Should -HaveCount 4
            }
            It 'contains the results for successful removals' {
                $actual[0].ComputerName | Should -Be $env:COMPUTERNAME
                $actual[0].Path | Should -Be $testFolder[0]
                $actual[0].Date | Should -Not -BeNullOrEmpty
                $actual[0].Exist | Should -BeFalse
                $actual[0].Action | Should -Be 'Remove'
                $actual[0].Error | Should -BeNullOrEmpty

                $actual[1].ComputerName | Should -Be $env:COMPUTERNAME
                $actual[1].Path | Should -Be $testFolder[1]
                $actual[1].Date | Should -Not -BeNullOrEmpty
                $actual[1].Exist | Should -BeFalse
                $actual[1].Action | Should -Be 'Remove'
                $actual[1].Error | Should -BeNullOrEmpty
            }
            It 'contains the results for failed removals' {
                $actual[2].ComputerName | Should -Be $env:COMPUTERNAME
                $actual[2].Path | Should -Be 'notExistingFolder'
                $actual[2].Date | Should -Not -BeNullOrEmpty
                $actual[2].Exist | Should -BeFalse
                $actual[2].Action | Should -BeNullOrEmpty
                $actual[2].Error | Should -Be 'Path not found'

                $actual[3].ComputerName | Should -Be $env:COMPUTERNAME
                $actual[3].Path | Should -Be $testFile
                $actual[3].Date | Should -Not -BeNullOrEmpty
                $actual[3].Exist | Should -BeTrue
                $actual[3].Action | Should -BeNullOrEmpty
                $actual[3].Error | Should -Be 'Path not a folder'
            }
        }
        It 'a summary mail is sent to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Bcc -eq $ScriptAdmin) -and
                ($Priority -eq 'High') -and
                ($Subject -eq 'Removed 2/4 folders, 2 removal errors') -and
                ($Attachments -like '*log.xlsx') -and
                ($Message -like '*
                *Successfully removed folders*2*
                *Errors while removing folders*2*
                *Imported Excel file rows*4*
                *Not existing folders*3*')
            }
        } -Tag test
    }
}