#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Path              = (New-Item 'TestDrive:/a' -ItemType File).FullName
        OlderThanUnit     = 'Year'
        OlderThanQuantity = 0
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
Describe 'remove' {
    BeforeAll {
        $test = @{
            Files   = @(
                "TestDrive:/b",
                "TestDrive:/c"
            ).ForEach(
                { New-Item $_ -ItemType File }
            )
            Folders = @(
                "TestDrive:/f1",
                "TestDrive:/f2"
            ).ForEach(
                { New-Item $_ -ItemType Directory }
            )
        }

        . $testScript @testParams
    }
    It 'the requested file' {
        $testParams.Path | Should -Not -Exist
    }
    Context 'do not remove' {
        It 'other files' {
            $test.Files.foreach(
                { $_.FullName | Should -Exist }
            )
        }
        It 'other folders' {
            $test.Folders.foreach(
                { $_.FullName | Should -Exist }
            )
        }
    }
}