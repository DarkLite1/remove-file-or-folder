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
Describe 'OlderThanUnit, OlderThanQuantity and OlderThanBasedOn are required for Type <_>' -ForEach @(
    'File', 'FilesInFolder'
) {
    It 'throws when they are missing' {
        { . $testScript -Type $_ -Path 'TestDrive:\' -OlderThanUnit 'Day' -OlderThanQuantity 0 } |
        Should -Throw "*Parameters 'OlderThanUnit', 'OlderThanQuantity' and 'OlderThanBasedOn' are mandatory for type '$_'*"
    }
}
Describe 'age validation' {
    It 'rejects negative quantity <_> without deleting the file' -ForEach @(-1, [int]::MinValue) {
        $testFile = New-Item "TestDrive:/negative_$_.txt" -ItemType File

        { . $testScript -Type File -Path $testFile.FullName -OlderThanUnit Day -OlderThanQuantity $_ -OlderThanBasedOn LastWriteTime } |
        Should -Throw '*OlderThanQuantity*'

        $testFile.FullName | Should -Exist
    }

    It 'continues to allow zero to remove a recent file' {
        $testFile = New-Item 'TestDrive:/zero.txt' -ItemType File

        $actual = . $testScript -Type File -Path $testFile.FullName -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn LastWriteTime

        $actual.Action | Should -Be 'Removed'
        $testFile.FullName | Should -Not -Exist
    }
}
Describe 'a path that does not exist is reported for Type <_>' -ForEach @(
    'File', 'FilesInFolder'
) {
    It 'with Error Path not found' {
        $actual = . $testScript -Type $_ -Path 'TestDrive:\notExisting' -OlderThanUnit 'Day' -OlderThanQuantity 0 -OlderThanBasedOn 'CreationTime'

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
            OlderThanBasedOn  = 'CreationTime'
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
            OlderThanBasedOn  = 'CreationTime'
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
                OlderThanBasedOn  = 'CreationTime'
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
    Context 'a subfolder that cannot be read' {
        BeforeAll {
            $testRoot = (New-Item 'TestDrive:/emptyUnreadable' -ItemType Directory).FullName
            $testDenied = (New-Item "$testRoot\denied" -ItemType Directory).FullName
            $testEmptyFolder = New-Item "$testRoot\readable\empty" -ItemType Directory -Force

            $testUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $testDenyRule = [System.Security.AccessControl.FileSystemAccessRule]::new(
                $testUser, 'ListDirectory', 'Deny'
            )
            $testAcl = Get-Acl -LiteralPath $testDenied
            $testAcl.AddAccessRule($testDenyRule)
            Set-Acl -LiteralPath $testDenied -AclObject $testAcl

            try {
                $actual = . $testScript -Type 'EmptyFolders' -Path $testRoot -WarningAction SilentlyContinue -ErrorVariable testErrors -ErrorAction SilentlyContinue
            }
            finally {
                $testAcl = Get-Acl -LiteralPath $testDenied
                $testAcl.RemoveAccessRule($testDenyRule) | Out-Null
                Set-Acl -LiteralPath $testDenied -AclObject $testAcl
            }
        }
        It 'does not stop the removal of other empty folders' {
            $testEmptyFolder.FullName | Should -Not -Exist
            ($actual | Where-Object FullName -EQ $testEmptyFolder.FullName).Action |
            Should -Be 'Removed'
        }
        It 'is reported once as an error' {
            @($actual | Where-Object FullName -EQ $testDenied) | Should -HaveCount 1
            ($actual | Where-Object FullName -EQ $testDenied).Error |
            Should -Not -BeNullOrEmpty
        }
        It 'writes no error to the error stream' {
            $testErrors | Should -BeNullOrEmpty
        }
    }
}
Describe 'ExcludeFolder' {
    BeforeEach {
        $testRoot = (New-Item "TestDrive:/exclude_$([guid]::NewGuid())" -ItemType Directory).FullName

        $testRemoveFile = New-Item "$testRoot\remove.txt" -ItemType File
        $testKeepFile = New-Item "$testRoot\Keep\PrintHistory.json" -ItemType File -Force
        $testKeepDeepFile = New-Item "$testRoot\Keep\sub\file.txt" -ItemType File -Force
        $testKeepEmptyFolder = New-Item "$testRoot\Keep\empty" -ItemType Directory
        $testRemoveEmptyFolder = New-Item "$testRoot\other\empty" -ItemType Directory -Force
        $testSimilarName = New-Item "$testRoot\Keeper\file.txt" -ItemType File -Force
    }
    Context 'Type FilesInFolder' {
        BeforeEach {
            $testParams = @{
                Type              = 'FilesInFolder'
                Path              = $testRoot
                ExcludeFolder     = @("$testRoot\keep")
                OlderThanUnit     = 'Day'
                OlderThanQuantity = 0
                OlderThanBasedOn  = 'CreationTime'
                Recurse           = $true
            }
        }
        It 'does not remove files in the excluded folder or its subfolders' {
            $actual = . $testScript @testParams

            $testKeepFile.FullName | Should -Exist
            $testKeepDeepFile.FullName | Should -Exist
            $actual.FullName | Should -Not -Contain $testKeepFile.FullName
        }
        It 'removes the other files' {
            . $testScript @testParams

            $testRemoveFile.FullName | Should -Not -Exist
        }
        It 'removes files in a folder that only starts with the same name' {
            . $testScript @testParams

            $testSimilarName.FullName | Should -Not -Exist
        }
        It 'accepts an excluded folder with a trailing backslash' {
            $testParams.ExcludeFolder = @("$testRoot\Keep\")

            . $testScript @testParams

            $testKeepFile.FullName | Should -Exist
        }
        It 'ignores read errors in the excluded folder' {
            $testUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $testDenyRule = [System.Security.AccessControl.FileSystemAccessRule]::new(
                $testUser, 'ListDirectory', 'Deny'
            )
            $testDenied = "$testRoot\Keep\sub"
            $testAcl = Get-Acl -LiteralPath $testDenied
            $testAcl.AddAccessRule($testDenyRule)
            Set-Acl -LiteralPath $testDenied -AclObject $testAcl

            try {
                $actual = . $testScript @testParams -WarningVariable testWarnings -WarningAction SilentlyContinue
            }
            finally {
                $testAcl = Get-Acl -LiteralPath $testDenied
                $testAcl.RemoveAccessRule($testDenyRule) | Out-Null
                Set-Acl -LiteralPath $testDenied -AclObject $testAcl
            }

            $actual.Error | Where-Object { $_ } | Should -BeNullOrEmpty
            $testWarnings | Should -BeNullOrEmpty
        }
    }
    Context 'Type EmptyFolders' {
        BeforeEach {
            $testParams = @{
                Type          = 'EmptyFolders'
                Path          = $testRoot
                ExcludeFolder = @("$testRoot\Keep")
            }
        }
        It 'does not remove empty folders in the excluded folder' {
            . $testScript @testParams

            $testKeepEmptyFolder.FullName | Should -Exist
        }
        It 'does not remove the excluded folder when it is empty' {
            $testEmptyExcluded = New-Item "$testRoot\EmptyKeep" -ItemType Directory
            $testParams.ExcludeFolder = @($testEmptyExcluded.FullName)

            . $testScript @testParams

            $testEmptyExcluded.FullName | Should -Exist
        }
        It 'removes the other empty folders' {
            . $testScript @testParams

            $testRemoveEmptyFolder.FullName | Should -Not -Exist
        }
    }
}
Describe 'ExcludeFile' {
    BeforeEach {
        $testRoot = (New-Item "TestDrive:/excludeFile_$([guid]::NewGuid())" -ItemType Directory).FullName

        $testKeepFile = New-Item "$testRoot\sub\PrintHistory.json" -ItemType File -Force
        $testRemoveFile = New-Item "$testRoot\sub\other.json" -ItemType File
        $testRemoveTopFile = New-Item "$testRoot\top.txt" -ItemType File

        $testParams = @{
            Type              = 'FilesInFolder'
            Path              = $testRoot
            OlderThanUnit     = 'Day'
            OlderThanQuantity = 0
            OlderThanBasedOn  = 'CreationTime'
            Recurse           = $true
            ExcludeFile       = @($testKeepFile.FullName.ToUpper())
        }

        $actual = . $testScript @testParams
    }
    It 'does not remove the excluded file, regardless of casing' {
        $testKeepFile.FullName | Should -Exist
        $actual.FullName | Should -Not -Contain $testKeepFile.FullName
    }
    It 'removes the other files' {
        $testRemoveFile.FullName | Should -Not -Exist
        $testRemoveTopFile.FullName | Should -Not -Exist
    }
}
Describe 'hidden items and retrieval errors' {
    It 'removes an explicitly targeted hidden file' {
        $testFile = New-Item 'TestDrive:/hidden-target.txt' -ItemType File
        $testFile.Attributes = $testFile.Attributes -bor [System.IO.FileAttributes]::Hidden

        $actual = . $testScript -Type File -Path $testFile.FullName -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime

        $testFile.FullName | Should -Not -Exist
        $actual.Action | Should -Be 'Removed'
    }

    It 'removes nested hidden empty folders but preserves hidden content' {
        $testRoot = (New-Item 'TestDrive:/hidden-folders' -ItemType Directory).FullName
        $testEmpty = New-Item "$testRoot/empty/child" -ItemType Directory -Force
        $testEmpty.Parent.Attributes = $testEmpty.Parent.Attributes -bor [System.IO.FileAttributes]::Hidden
        $testKeep = New-Item "$testRoot/keep/hidden.txt" -ItemType File -Force
        $testKeep.Attributes = $testKeep.Attributes -bor [System.IO.FileAttributes]::Hidden
        $testKeep.Directory.Attributes = $testKeep.Directory.Attributes -bor [System.IO.FileAttributes]::Hidden

        . $testScript -Type EmptyFolders -Path $testRoot

        "$testRoot/empty" | Should -Not -Exist
        $testKeep.FullName | Should -Exist
        $testRoot | Should -Exist
    }

    It 'returns a retrieval error when a file becomes unreadable after validation' {
        $testFile = New-Item 'TestDrive:/read-error.txt' -ItemType File
        Mock Get-Item { Write-Error 'Read failed' -TargetObject $LiteralPath } -ParameterFilter {
            $LiteralPath -eq $testFile.FullName
        }

        $actual = @(. $testScript -Type File -Path $testFile.FullName -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime -WarningAction SilentlyContinue)

        $actual | Should -HaveCount 1
        $actual[0].FullName | Should -Be $testFile.FullName
        $actual[0].Error | Should -BeLike '*Read failed*'
        $actual[0].Action | Should -BeNullOrEmpty
        $testFile.FullName | Should -Exist
    }
}
Describe 'normalized exclusions' {
    It 'protects a folder using <Suffix> for <Type>' -ForEach @(
        @{ Suffix = 'Keep\.'; Type = 'FilesInFolder' }
        @{ Suffix = 'Keep/sub/..'; Type = 'FilesInFolder' }
        @{ Suffix = 'Keep\.'; Type = 'EmptyFolders' }
        @{ Suffix = 'Keep/sub/..'; Type = 'EmptyFolders' }
    ) {
        $testRoot = (New-Item "TestDrive:/normalized_$([guid]::NewGuid())" -ItemType Directory).FullName
        $testKeep = New-Item "$testRoot/Keep/empty" -ItemType Directory -Force
        $testFile = New-Item "$testRoot/Keep/protected.txt" -ItemType File
        $testOther = New-Item "$testRoot/Other/empty" -ItemType Directory -Force
        $testParams = @{
            Type = $Type
            Path = $testRoot
            ExcludeFolder = @("$testRoot/$Suffix")
            OlderThanUnit = 'Day'
            OlderThanQuantity = 0
            OlderThanBasedOn = 'CreationTime'
            Recurse = $true
        }

        . $testScript @testParams

        $testKeep.FullName | Should -Exist
        $testFile.FullName | Should -Exist
        if ($Type -eq 'EmptyFolders') { $testOther.FullName | Should -Not -Exist }
    }

    It 'protects a file using <_>' -ForEach @(
        'Keep/protected.txt', 'Keep\sub\..\protected.txt', 'Keep\.\protected.txt'
    ) {
        $testRoot = (New-Item "TestDrive:/normalizedFile_$([guid]::NewGuid())" -ItemType Directory).FullName
        $testFile = New-Item "$testRoot/Keep/protected.txt" -ItemType File -Force
        $testOther = New-Item "$testRoot/remove.txt" -ItemType File

        . $testScript -Type FilesInFolder -Path $testRoot -ExcludeFile "$testRoot/$_" -Recurse $true -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime

        $testFile.FullName | Should -Exist
        $testOther.FullName | Should -Not -Exist
    }

    It 'resolves relative exclusions in the worker location' {
        $testRoot = (New-Item 'TestDrive:/relative' -ItemType Directory).FullName
        $testFile = New-Item "$testRoot/Keep/protected.txt" -ItemType File -Force
        Push-Location $testRoot
        try {
            . $testScript -Type FilesInFolder -Path . -ExcludeFolder './Keep' -Recurse $true -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime
        }
        finally { Pop-Location }

        $testFile.FullName | Should -Exist
    }
}
Describe 'OlderThanBasedOn <BasedOn>' -ForEach @(
    @{ BasedOn = 'CreationTime'; RemovesDailyWrittenFile = $true; RemovesCopiedFile = $false }
    @{ BasedOn = 'LastWriteTime'; RemovesDailyWrittenFile = $false; RemovesCopiedFile = $true }
) {
    BeforeAll {
        $testRoot = (New-Item "TestDrive:/basedOn_$BasedOn" -ItemType Directory).FullName

        # created long ago, but written every day, like a history file
        $testDailyWrittenFile = New-Item "$testRoot\history.json" -ItemType File
        $testDailyWrittenFile.CreationTime = (Get-Date).AddDays(-100)
        $testDailyWrittenFile.LastWriteTime = Get-Date

        # copied today, but with its original old content date
        $testCopiedFile = New-Item "$testRoot\copied.txt" -ItemType File
        $testCopiedFile.CreationTime = Get-Date
        $testCopiedFile.LastWriteTime = (Get-Date).AddDays(-100)

        $testParams = @{
            Type              = 'FilesInFolder'
            Path              = $testRoot
            OlderThanUnit     = 'Day'
            OlderThanQuantity = 30
            OlderThanBasedOn  = $BasedOn
            Recurse           = $false
        }

        $actual = . $testScript @testParams
    }
    It 'a file created long ago but written today is removed: <RemovesDailyWrittenFile>' {
        Test-Path -LiteralPath $testDailyWrittenFile.FullName |
        Should -Be (-not $RemovesDailyWrittenFile)
    }
    It 'a file created today with an old last write time is removed: <RemovesCopiedFile>' {
        Test-Path -LiteralPath $testCopiedFile.FullName |
        Should -Be (-not $RemovesCopiedFile)
    }
    It 'reports both dates of a removed file' {
        $actual | Should -Not -BeNullOrEmpty
        $actual | ForEach-Object {
            $_.CreationTime | Should -Not -BeNullOrEmpty
            $_.LastWriteTime | Should -Not -BeNullOrEmpty
        }
    }
}
