#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Path = (New-Item 'TestDrive:/f' -ItemType Directory).FullName
    }
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'Path'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'remove folders' {
    BeforeAll {
        @(
            "$($testParams.Path)/EmptyFolders/a/1/2/3",
            "$($testParams.Path)/EmptyFolders/b/1/2",
            "$($testParams.Path)/EmptyFolders/c/1"
        ).ForEach(
            { New-Item $_ -ItemType Directory }
        )

        New-Item "$($testParams.Path)/Folder" -ItemType Directory
        $testFile = New-Item "$($testParams.Path)/Folder\a.txt" -ItemType File

        . $testScript @testParams
    }
    It 'when they are empty' {
        "$($testParams.Path)/EmptyFolders" | Should -Not -Exist
    }
    Context 'do not remove' {
        It 'the parent folder' {
            $testParams.Path | Should -Exist
        }
        It 'folders that are not empty' {
            $testFile | Should -Exist
        }
    }
}