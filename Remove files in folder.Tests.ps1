#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Path              = (New-Item 'TestDrive:/a' -ItemType Directory).FullName
        OlderThanUnit     = 'Month'
        OlderThanQuantity = 1
        Recurse           = $false
    }
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'Path', 'OlderThanUnit', 'OlderThanQuantity'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'a file is' {
    Context 'not removed when it is created more recently than' {
        BeforeAll {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.OlderThanQuantity = 3

            $testFile = New-Item -Path "$($testNewParams.Path)\file.txt" -ItemType File
        }
        AfterEach {
            . $testScript @testNewParams

            $testFile | Should -Exist
        }
        It 'Day' {
            $testNewParams.OlderThanUnit = 'Day'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddDays(-2)
            }
        }
        It 'Month' {
            $testNewParams.OlderThanUnit = 'Month'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddMonths(-2)
            }
        }
        It 'Year' {
            $testNewParams.OlderThanUnit = 'Year'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddYears(-2)
            }
        }
    }
    Context 'removed when it is OlderThan' {
        BeforeEach {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.OlderThanQuantity = 3

            $testFile = New-Item -Path "$($testNewParams.Path)\file.txt" -ItemType File -Force
        }
        AfterEach {
            . $testScript @testNewParams

            $testFile | Should -Not -Exist
        }
        It 'Day' {
            $testNewParams.OlderThanUnit = 'Day'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddDays(-4)
            }
        }
        It 'Month' {
            $testNewParams.OlderThanUnit = 'Month'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddMonths(-4)
            }
        }
        It 'Year' {
            $testNewParams.OlderThanUnit = 'Year'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddYears(-4)
            }
        }
    }
}