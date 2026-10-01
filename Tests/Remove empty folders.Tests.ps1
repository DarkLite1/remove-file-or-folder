#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = Join-Path (Split-Path $PSScriptRoot) (
        (Split-Path $PSCommandPath -Leaf).Replace('.Tests.ps1', '.ps1')
    )
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
    It 'when they are empty and read-only' {
        $testFolder = New-Item "$($testParams.Path)/ReadOnly/a" -ItemType Directory
        $testFolder.Attributes = $testFolder.Attributes -bor [System.IO.FileAttributes]::ReadOnly

        . $testScript @testParams

        "$($testParams.Path)/ReadOnly" | Should -Not -Exist
    }
    Context 'do not remove' {
        It 'the parent folder' {
            $testParams.Path | Should -Exist
        }
        It 'folders that are not empty' {
            $testFile | Should -Exist
        }
        It 'folders that only contain a hidden file' {
            $testHiddenFile = New-Item "$($testParams.Path)/Hidden/h.txt" -ItemType File -Force
            $testHiddenFile.Attributes = 'Hidden'

            . $testScript @testParams

            $testHiddenFile.FullName | Should -Exist
        }
    }
}
Describe 'a folder that is no longer empty when it is removed' {
    It 'is not removed and its content is kept' {
        $testFolder = New-Item "$($testParams.Path)/Race" -ItemType Directory
        $testFile = New-Item "$($testParams.Path)/Race/late.txt" -ItemType File

        # simulates a file arriving between finding and removing the folder
        $testScriptText = (Get-Content -LiteralPath $testScript -Raw).Replace(
            'Write-Verbose "Remove empty folder ''$emptyFolder''"',
            'if ($emptyFolder.Name -eq ''Race'') { New-Item -Path (Join-Path $emptyFolder.FullName ''late.txt'') -ItemType File -Force | Out-Null }'
        )
        $testScriptText | Should -BeLike '*late.txt*'
        Remove-Item -LiteralPath $testFile.FullName

        $actual = & ([scriptblock]::Create($testScriptText)) -Path $testParams.Path

        $testFile.FullName | Should -Exist
        ($actual | Where-Object FullName -EQ $testFolder.FullName).Error |
        Should -Not -BeNullOrEmpty
    }
}