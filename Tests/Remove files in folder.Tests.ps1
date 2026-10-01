#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = Join-Path (Split-Path $PSScriptRoot) (
        (Split-Path $PSCommandPath -Leaf).Replace('.Tests.ps1', '.ps1')
    )
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
Describe 'a file that cannot be removed' {
    BeforeAll {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.OlderThanQuantity = 0

        $testFile = New-Item -Path "$($testNewParams.Path)\locked.txt" -ItemType File -Force
        $testLock = [System.IO.File]::Open(
            $testFile.FullName, 'Open', 'Read', 'None'
        )

        try {
            $actual = . $testScript @testNewParams -WarningVariable testWarnings -WarningAction SilentlyContinue
        }
        finally {
            $testLock.Dispose()
        }
    }
    It 'is reported with the file path, not the folder path' {
        ($actual | Where-Object FullName -EQ $testFile.FullName).Error |
        Should -Not -BeNullOrEmpty
        $testWarnings | Should -BeLike "*Failed to remove file '$($testFile.FullName)'*"
    }
}