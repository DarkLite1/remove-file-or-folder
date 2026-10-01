#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = Join-Path (Split-Path $PSScriptRoot) (
        (Split-Path $PSCommandPath -Leaf).Replace('.Tests.ps1', '.ps1')
    )
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('Type', 'Path') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'OlderThanUnit and OlderThanQuantity are required for Type <_>' -ForEach @(
    'File', 'FilesInFolder'
) {
    It 'throws when they are missing' {
        { . $testScript -Type $_ -Path 'TestDrive:\' } |
        Should -Throw "*Parameters 'OlderThanUnit' and 'OlderThanQuantity' are mandatory for type '$_'*"
    }
}
Describe 'a path that does not exist is reported for Type <_>' -ForEach @(
    'File', 'FilesInFolder'
) {
    It 'with Error Path not found' {
        $actual = . $testScript -Type $_ -Path 'TestDrive:\notExisting' -OlderThanUnit 'Day' -OlderThanQuantity 0

        $actual.Type | Should -Be $_
        $actual.Error | Should -Be 'Path not found'
    }
}
Describe 'Type File' {
    BeforeAll {
        $testParams = @{
            Type              = 'File'
            Path              = (New-Item 'TestDrive:/File/a' -ItemType File -Force).FullName
            OlderThanUnit     = 'Year'
            OlderThanQuantity = 0
        }

        $test = @{
            Files   = @(
                'TestDrive:/File/b',
                'TestDrive:/File/c'
            ).ForEach(
                { New-Item $_ -ItemType File }
            )
            Folders = @(
                'TestDrive:/File/f1',
                'TestDrive:/File/f2'
            ).ForEach(
                { New-Item $_ -ItemType Directory }
            )
        }

        $actual = . $testScript @testParams
    }
    It 'removes the requested file' {
        $testParams.Path | Should -Not -Exist
        $actual.Action | Should -Be 'Removed'
    }
    Context 'does not remove' {
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
Describe 'Type FilesInFolder' {
    BeforeAll {
        $testParams = @{
            Type              = 'FilesInFolder'
            Path              = (New-Item 'TestDrive:/FilesInFolder' -ItemType Directory).FullName
            OlderThanUnit     = 'Month'
            OlderThanQuantity = 3
            Recurse           = $false
        }
    }
    Context 'a file is not removed when it is created more recently than' {
        BeforeAll {
            $testFile = New-Item -Path "$($testParams.Path)\file.txt" -ItemType File
        }
        AfterEach {
            $testNewParams = $testParams.Clone()
            $testNewParams.OlderThanUnit = $testUnit

            . $testScript @testNewParams

            $testFile | Should -Exist
        }
        It 'Day' {
            $testUnit = 'Day'
            $testFile.CreationTime = (Get-Date).AddDays(-2)
        }
        It 'Month' {
            $testUnit = 'Month'
            $testFile.CreationTime = (Get-Date).AddMonths(-2)
        }
        It 'Year' {
            $testUnit = 'Year'
            $testFile.CreationTime = (Get-Date).AddYears(-2)
        }
    }
    Context 'a file is removed when it is OlderThan' {
        BeforeEach {
            $testFile = New-Item -Path "$($testParams.Path)\file.txt" -ItemType File -Force
        }
        AfterEach {
            $testNewParams = $testParams.Clone()
            $testNewParams.OlderThanUnit = $testUnit

            . $testScript @testNewParams

            $testFile | Should -Not -Exist
        }
        It 'Day' {
            $testUnit = 'Day'
            $testFile.CreationTime = (Get-Date).AddDays(-4)
        }
        It 'Month' {
            $testUnit = 'Month'
            $testFile.CreationTime = (Get-Date).AddMonths(-4)
        }
        It 'Year' {
            $testUnit = 'Year'
            $testFile.CreationTime = (Get-Date).AddYears(-4)
        }
    }
    Context 'Recurse' {
        BeforeEach {
            $testFile = New-Item -Path "$($testParams.Path)\sub\file.txt" -ItemType File -Force
        }
        It 'false does not remove files in subfolders' {
            $testNewParams = $testParams.Clone()
            $testNewParams.OlderThanQuantity = 0

            . $testScript @testNewParams

            $testFile | Should -Exist
        }
        It 'true removes files in subfolders' {
            $testNewParams = $testParams.Clone()
            $testNewParams.OlderThanQuantity = 0
            $testNewParams.Recurse = $true

            . $testScript @testNewParams

            $testFile | Should -Not -Exist
        }
    }
    Context 'a file that cannot be removed' {
        BeforeAll {
            $testNewParams = $testParams.Clone()
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
    Context 'a subfolder that cannot be read' {
        BeforeAll {
            $testRoot = (New-Item 'TestDrive:/unreadable' -ItemType Directory).FullName
            $testDenied = (New-Item "$testRoot\denied" -ItemType Directory).FullName
            $testFile = New-Item "$testRoot\readable\file.txt" -ItemType File -Force

            $testUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $testDenyRule = [System.Security.AccessControl.FileSystemAccessRule]::new(
                $testUser, 'ListDirectory', 'Deny'
            )
            $testAcl = Get-Acl -LiteralPath $testDenied
            $testAcl.AddAccessRule($testDenyRule)
            Set-Acl -LiteralPath $testDenied -AclObject $testAcl

            $testNewParams = @{
                Type              = 'FilesInFolder'
                Path              = $testRoot
                OlderThanUnit     = 'Day'
                OlderThanQuantity = 0
                Recurse           = $true
            }

            try {
                $actual = . $testScript @testNewParams -WarningAction SilentlyContinue
            }
            finally {
                $testAcl = Get-Acl -LiteralPath $testDenied
                $testAcl.RemoveAccessRule($testDenyRule) | Out-Null
                Set-Acl -LiteralPath $testDenied -AclObject $testAcl
            }
        }
        It 'does not stop the removal of other files' {
            $testFile.FullName | Should -Not -Exist
        }
        It 'is reported as an error' {
            ($actual | Where-Object FullName -EQ $testDenied).Error |
            Should -Not -BeNullOrEmpty
        }
    }
}
Describe 'Type EmptyFolders' {
    BeforeAll {
        $testParams = @{
            Type = 'EmptyFolders'
            Path = (New-Item 'TestDrive:/EmptyFolders' -ItemType Directory).FullName
        }

        @(
            "$($testParams.Path)/Empty/a/1/2/3",
            "$($testParams.Path)/Empty/b/1/2",
            "$($testParams.Path)/Empty/c/1"
        ).ForEach(
            { New-Item $_ -ItemType Directory }
        )

        $testFile = New-Item "$($testParams.Path)/Folder/a.txt" -ItemType File -Force

        . $testScript @testParams
    }
    It 'removes folders when they are empty' {
        "$($testParams.Path)/Empty" | Should -Not -Exist
    }
    It 'removes folders when they are empty and read-only' {
        $testFolder = New-Item "$($testParams.Path)/ReadOnly/a" -ItemType Directory
        $testFolder.Attributes = $testFolder.Attributes -bor [System.IO.FileAttributes]::ReadOnly

        . $testScript @testParams

        "$($testParams.Path)/ReadOnly" | Should -Not -Exist
    }
    Context 'does not remove' {
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
    Context 'a folder that is no longer empty when it is removed' {
        It 'is not removed and its content is kept' {
            $testFolder = New-Item "$($testParams.Path)/Race" -ItemType Directory
            $testFile = Join-Path $testFolder.FullName 'late.txt'

            # simulates a file arriving between finding and removing the folder
            $testScriptText = (Get-Content -LiteralPath $testScript -Raw).Replace(
                'Write-Verbose "Remove empty folder ''$emptyFolder''"',
                'if ($emptyFolder.Name -eq ''Race'') { New-Item -Path (Join-Path $emptyFolder.FullName ''late.txt'') -ItemType File -Force | Out-Null }'
            )
            $testScriptText | Should -BeLike '*late.txt*'

            $actual = & ([scriptblock]::Create($testScriptText)) @testParams

            $testFile | Should -Exist
            ($actual | Where-Object FullName -EQ $testFolder.FullName).Error |
            Should -Not -BeNullOrEmpty
        }
    }
}
