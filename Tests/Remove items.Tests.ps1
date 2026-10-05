#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = Join-Path (Split-Path $PSScriptRoot) (
        (Split-Path $PSCommandPath -Leaf).Replace('.Tests.ps1', '.ps1')
    )
    $testTokens = $null
    $testParseErrors = $null
    $testAst = [System.Management.Automation.Language.Parser]::ParseFile($testScript, [ref]$testTokens, [ref]$testParseErrors)
    $testParseErrors | Should -BeNullOrEmpty
    foreach ($testFunctionName in 'Get-ExclusiveCutoffHC', 'New-ReadErrorResultHC') {
        $testFunction = $testAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq $testFunctionName
            }, $true)
        . ([scriptblock]::Create($testFunction.Extent.Text))
    }
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
Describe 'worker diagnostic messages' {
    It 'does not invoke Write-Verbose for a quiet run of 1000 files' {
        $testRoot = (New-Item 'TestDrive:/quiet-files' -ItemType Directory).FullName
        foreach ($fileIndex in 1..1000) {
            [System.IO.File]::WriteAllText((Join-Path $testRoot "$fileIndex.txt"), '')
        }
        Mock Write-Verbose

        $actual = @(& $testScript -Type FilesInFolder -Path $testRoot -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn LastWriteTime -Verbose:$false)

        $actual | Should -HaveCount 1000
        @($actual | Where-Object Action -EQ Removed) | Should -HaveCount 1000
        Should -Not -Invoke Write-Verbose -Scope It
    }

    It 'shows the exact calendar cutoff and logs only completed file removals' {
        $testRoot = (New-Item 'TestDrive:/verbose-files' -ItemType Directory).FullName
        $testOld = New-Item "$testRoot/old.txt" -ItemType File
        $testOld.LastWriteTime = [datetime]'2026-10-01T23:59:59'
        $testKeep = New-Item "$testRoot/keep.txt" -ItemType File
        $testKeep.LastWriteTime = [datetime]'2026-10-02T00:00:00'
        Mock Get-Date { [datetime]'2026-10-02T12:00:00' }

        $actual = @(& $testScript -Type FilesInFolder -Path $testRoot -OlderThanUnit Day -OlderThanQuantity 1 -OlderThanBasedOn LastWriteTime -Verbose 4>&1)
        $messages = @($actual | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] } | ForEach-Object Message)

        $messages | Should -HaveCount 2
        $messages[0] | Should -BeLike '*LastWriteTime before 2026-10-02 00:00:00 (exclusive calendar cutoff, local time)*'
        $messages[1] | Should -Be "Removed file '$($testOld.FullName)'"
        $testOld.FullName | Should -Not -Exist
        $testKeep.FullName | Should -Exist
    }

    It 'describes quantity zero as disabled age filtering' {
        $testFile = New-Item 'TestDrive:/verbose-zero.txt' -ItemType File
        $actual = @(& $testScript -Type File -Path $testFile.FullName -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime -Verbose 4>&1)
        $messages = @($actual | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] } | ForEach-Object Message)

        $messages[0] | Should -Be "Age filtering disabled for '$($testFile.FullName)'; exclusions and Recurse still apply"
        $messages[1] | Should -Be "Removed file '$($testFile.FullName)'"
    }

    It 'does not invoke Write-Verbose for quiet empty-folder cleanup' {
        $testRoot = (New-Item 'TestDrive:/quiet-folders/empty' -ItemType Directory -Force).Parent.FullName
        Mock Write-Verbose

        $actual = @(& $testScript -Type EmptyFolders -Path $testRoot -Verbose:$false)

        $actual | Should -HaveCount 1
        $actual[0].Action | Should -Be Removed
        Should -Not -Invoke Write-Verbose -Scope It
    }
}
Describe 'attribute exclusions' {
    It 'protects <Attribute> items and prunes their folder trees for <Type>' -ForEach @(
        foreach ($attribute in 'Hidden', 'System') {
            foreach ($type in 'FilesInFolder', 'EmptyFolders') {
                @{ Attribute = $attribute; Type = $type }
            }
        }
    ) {
        $testRoot = (New-Item "TestDrive:/attributes-$Attribute-$Type" -ItemType Directory).FullName
        $testProtected = New-Item "$testRoot/protected" -ItemType Directory
        $testNested = New-Item "$testRoot/protected/child.txt" -ItemType File
        $testHiddenFile = New-Item "$testRoot/protected.txt" -ItemType File
        $testOrdinary = New-Item "$testRoot/ordinary.txt" -ItemType File
        $testEmpty = New-Item "$testRoot/empty" -ItemType Directory
        $testProtected.Attributes = $testProtected.Attributes -bor [System.IO.FileAttributes]$Attribute
        $testHiddenFile.Attributes = $testHiddenFile.Attributes -bor [System.IO.FileAttributes]$Attribute
        $testGetChildItem = Get-Command Get-ChildItem -CommandType Cmdlet
        Mock Get-ChildItem {
            $parameters = @{ LiteralPath = $LiteralPath; Force = $true }
            if ($PesterBoundParameters.Directory) { $parameters.Directory = $true }
            & $testGetChildItem @parameters
        }

        $actual = @(& $testScript -Type $Type -Path $testRoot -ExcludeAttributes @($Attribute) -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn LastWriteTime -Recurse $true)

        $testProtected.FullName | Should -Exist
        $testNested.FullName | Should -Exist
        $testHiddenFile.FullName | Should -Exist
        @($actual | Where-Object Error) | Should -HaveCount 0
        Should -Not -Invoke Get-ChildItem -Scope It -ParameterFilter { $LiteralPath -eq $testProtected.FullName }
        if ($Type -eq 'FilesInFolder') { $testOrdinary.FullName | Should -Not -Exist }
        else { $testEmpty.FullName | Should -Not -Exist }
    }
    It 'protects an explicitly selected <_> file' -ForEach @('Hidden', 'System') {
        $testFile = New-Item "TestDrive:/explicit-$_.txt" -ItemType File
        $testFile.Attributes = $testFile.Attributes -bor [System.IO.FileAttributes]$_

        $actual = @(& $testScript -Type File -Path $testFile.FullName -ExcludeAttributes @($_) -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn LastWriteTime)

        $actual | Should -HaveCount 0
        $testFile.FullName | Should -Exist
    }
    It 'skips a matching root for <_>' -ForEach @('FilesInFolder', 'EmptyFolders') {
        $testRoot = New-Item "TestDrive:/excluded-root-$_" -ItemType Directory
        $testFile = New-Item "$($testRoot.FullName)/keep.txt" -ItemType File
        $testEmpty = New-Item "$($testRoot.FullName)/empty" -ItemType Directory
        $testRoot.Attributes = $testRoot.Attributes -bor [System.IO.FileAttributes]::System
        Mock Get-ChildItem { throw 'Excluded root must not be enumerated' }

        $actual = @(& $testScript -Type $_ -Path $testRoot.FullName -ExcludeAttributes @('Hidden', 'System') -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn LastWriteTime -Recurse $true)

        $actual | Should -HaveCount 0
        $testFile.FullName | Should -Exist
        $testEmpty.FullName | Should -Exist
        Should -Not -Invoke Get-ChildItem -Scope It
    }
    It 'keeps a folder nonempty when only an excluded file remains' {
        $testRoot = (New-Item 'TestDrive:/nonempty-attributes/child' -ItemType Directory -Force).Parent.FullName
        $testFile = New-Item "$testRoot/child/keep.txt" -ItemType File
        $testFile.Attributes = $testFile.Attributes -bor [System.IO.FileAttributes]::Hidden

        $actual = @(& $testScript -Type EmptyFolders -Path $testRoot -ExcludeAttributes @('Hidden', 'System'))

        $actual | Should -HaveCount 0
        $testFile.FullName | Should -Exist
    }
    It 'continues to remove system files with an empty exclusion list' {
        $testFile = New-Item 'TestDrive:/default-system.txt' -ItemType File
        $testFile.Attributes = $testFile.Attributes -bor [System.IO.FileAttributes]::System

        $actual = @(& $testScript -Type File -Path $testFile.FullName -ExcludeAttributes @() -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn LastWriteTime)

        $actual[0].Action | Should -Be Removed
        $testFile.FullName | Should -Not -Exist
    }
    It 'rechecks attributes immediately before deleting <Type>' -ForEach @(
        @{ Type = 'File'; Anchor = '$fileToRemove.Refresh()'; Target = '$fileToRemove.FullName' }
        @{ Type = 'EmptyFolders'; Anchor = '$emptyFolder.Refresh()'; Target = '$emptyFolder.FullName' }
    ) {
        $testRoot = (New-Item "TestDrive:/attribute-race-$Type" -ItemType Directory).FullName
        $testItem = if ($Type -eq 'File') {
            New-Item "$testRoot/item.txt" -ItemType File
        }
        else { New-Item "$testRoot/empty" -ItemType Directory }
        $testPath = if ($Type -eq 'File') { $testItem.FullName } else { $testRoot }
        $testScriptText = (Get-Content -LiteralPath $testScript -Raw).Replace(
            $Anchor,
            "[System.IO.File]::SetAttributes($Target, [System.IO.File]::GetAttributes($Target) -bor [System.IO.FileAttributes]::Hidden); $Anchor"
        )
        $testScriptText | Should -Not -Be (Get-Content -LiteralPath $testScript -Raw)

        $actual = @(& ([scriptblock]::Create($testScriptText)) -Type $Type -Path $testPath -ExcludeAttributes @('Hidden') -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn LastWriteTime)

        $testItem.FullName | Should -Exist
        $actual | Should -HaveCount 0
    }
    It 'still reports non-excluded access failures for <_>' -ForEach @('FilesInFolder', 'EmptyFolders') {
        $testRoot = (New-Item "TestDrive:/attribute-denied-$_" -ItemType Directory).FullName
        $testDenied = (New-Item "$testRoot/denied" -ItemType Directory).FullName
        $testUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        $testRule = [System.Security.AccessControl.FileSystemAccessRule]::new($testUser, 'ListDirectory', 'Deny')
        $testAcl = Get-Acl -LiteralPath $testDenied
        $testAcl.AddAccessRule($testRule)
        Set-Acl -LiteralPath $testDenied -AclObject $testAcl
        try {
            $actual = @(& $testScript -Type $_ -Path $testRoot -ExcludeAttributes @('Hidden', 'System') -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn LastWriteTime -Recurse $true -WarningAction SilentlyContinue)
        }
        finally {
            $testAcl.RemoveAccessRule($testRule) | Out-Null
            Set-Acl -LiteralPath $testDenied -AclObject $testAcl
        }

        $actual | Should -HaveCount 1
        $actual[0].FullName | Should -Be $testDenied
        $actual[0].Error | Should -Not -BeNullOrEmpty
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
                '$emptyFolder.Delete()',
                'if ($emptyFolder.Name -eq ''Race'') { New-Item -Path (Join-Path $emptyFolder.FullName ''late.txt'') -ItemType File -Force | Out-Null }; $emptyFolder.Delete()'
            )
            $testScriptText | Should -BeLike '*late.txt*'

            $testWarnings = @()
            $actual = & ([scriptblock]::Create($testScriptText)) @testParams -WarningVariable testWarnings -WarningAction SilentlyContinue

            $testFile | Should -Exist
            ($actual | Where-Object FullName -EQ $testFolder.FullName).Error |
            Should -Not -BeNullOrEmpty
            ($testWarnings -join ' ') | Should -BeLike "*Failed to remove empty folder '$($testFolder.FullName)'*"
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
Describe 'streamed file processing' {
    It 'removes read-only hidden files for <_>' -ForEach @('File', 'FilesInFolder') {
        $testRoot = (New-Item "TestDrive:/readonly_$([guid]::NewGuid())" -ItemType Directory).FullName
        $testFile = New-Item "$testRoot/readonly.txt" -ItemType File
        $testFile.Attributes = [System.IO.FileAttributes]::ReadOnly -bor [System.IO.FileAttributes]::Hidden
        $testPath = if ($_ -eq 'File') { $testFile.FullName } else { $testRoot }

        $actual = . $testScript -Type $_ -Path $testPath -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime

        $actual.Action | Should -Be 'Removed'
        $testFile.FullName | Should -Not -Exist
    }

    It 'reports a file that disappears after enumeration' {
        $testRoot = (New-Item 'TestDrive:/disappearing' -ItemType Directory).FullName
        $testFile = New-Item "$testRoot/gone.txt" -ItemType File
        Mock Get-ChildItem {
            Remove-Item -LiteralPath $testFile.FullName
            $testFile
        }

        $actual = @(. $testScript -Type FilesInFolder -Path $testRoot -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime -WarningAction SilentlyContinue)

        $actual | Should -HaveCount 1
        $actual[0].FullName | Should -Be $testFile.FullName
        $actual[0].Error | Should -Not -BeNullOrEmpty
        $actual[0].Action | Should -BeNullOrEmpty
    }

    It 'preserves a file updated after its cached timestamp was selected' {
        $testRoot = (New-Item 'TestDrive:/updated-during-cleanup' -ItemType Directory).FullName
        $testFile = New-Item "$testRoot/active.txt" -ItemType File
        $testFile.LastWriteTime = [datetime]::Now.AddDays(-100)
        $testScriptText = (Get-Content -LiteralPath $testScript -Raw).Replace(
            '$fileToRemove.Refresh()',
            '[System.IO.File]::SetLastWriteTime($fileToRemove.FullName, [datetime]::Now); $fileToRemove.Refresh()'
        )
        $testScriptText | Should -BeLike '*SetLastWriteTime*'

        $testVerbose = @()
        $actual = @(& ([scriptblock]::Create($testScriptText)) -Type FilesInFolder -Path $testRoot -OlderThanUnit Day -OlderThanQuantity 30 -OlderThanBasedOn LastWriteTime -Verbose 4>&1)
        $testVerbose = @($actual | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] })
        $actual = @($actual | Where-Object { $_ -isnot [System.Management.Automation.VerboseRecord] })

        $actual | Should -HaveCount 0
        $testFile.FullName | Should -Exist
        @($testVerbose | Where-Object { $_.Message -like 'Removed file*' }) | Should -HaveCount 0
    }

    It 'removes a file before enumeration produces the next file' {
        $testRoot = (New-Item 'TestDrive:/streamed' -ItemType Directory).FullName
        $testFirst = New-Item "$testRoot/first.txt" -ItemType File
        $testSecond = New-Item "$testRoot/second.txt" -ItemType File
        Mock Get-ChildItem {
            $testFirst
            $testFirst.FullName | Should -Not -Exist
            $testSecond
        }

        $actual = @(. $testScript -Type FilesInFolder -Path $testRoot -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime)

        $actual | Should -HaveCount 2
        $testSecond.FullName | Should -Not -Exist
    }

    It 'handles duplicate case-insensitive exclusions and similar file names' {
        $testRoot = (New-Item 'TestDrive:/hash-exclusions' -ItemType Directory).FullName
        $testKeep = New-Item "$testRoot/state.txt" -ItemType File
        $testRemove = New-Item "$testRoot/state.txt.old" -ItemType File

        . $testScript -Type FilesInFolder -Path $testRoot -ExcludeFile @($testKeep.FullName, $testKeep.FullName.ToUpperInvariant()) -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime

        $testKeep.FullName | Should -Exist
        $testRemove.FullName | Should -Not -Exist
    }
}
Describe 'exclusive cutoff helper' {
    It 'returns the exclusive <Unit> boundary for quantity <Quantity>' -ForEach @(
        @{ Unit = 'Day'; Quantity = 30; ReferenceDate = '2026-10-02T15:30:00'; Expected = '2026-09-03' }
        @{ Unit = 'Day'; Quantity = 1; ReferenceDate = '2024-03-01T15:30:00'; Expected = '2024-03-01' }
        @{ Unit = 'Month'; Quantity = 1; ReferenceDate = '2024-03-31T15:30:00'; Expected = '2024-03-01' }
        @{ Unit = 'Month'; Quantity = 13; ReferenceDate = '2026-01-31T15:30:00'; Expected = '2025-01-01' }
        @{ Unit = 'Year'; Quantity = 1; ReferenceDate = '2024-02-29T15:30:00'; Expected = '2024-01-01' }
        @{ Unit = 'Year'; Quantity = 3; ReferenceDate = '2026-10-02T15:30:00'; Expected = '2024-01-01' }
    ) {
        $actual = Get-ExclusiveCutoffHC -ReferenceDate ([datetime]$ReferenceDate) -Unit $Unit -Quantity $Quantity

        $actual | Should -BeOfType ([datetime])
        $actual | Should -Be ([datetime]$Expected)
    }

    It 'rejects nonpositive quantity <_>' -ForEach @(0, -1) {
        { Get-ExclusiveCutoffHC -ReferenceDate ([datetime]'2026-10-02') -Unit Day -Quantity $_ } |
        Should -Throw '*Quantity*'
    }

    It 'preserves the overflow error for <_>' -ForEach @('Day', 'Month', 'Year') {
        { Get-ExclusiveCutoffHC -ReferenceDate ([datetime]'2026-10-02') -Unit $_ -Quantity ([int]::MaxValue) } |
        Should -Throw '*Invalid retention period*'
    }
}
Describe 'read-error result helper' {
    It 'preserves the result schema for <_>' -ForEach @('File', 'FilesInFolder', 'EmptyFolders') {
        $actual = @(New-ReadErrorResultHC -FullName 'z:\folder\[item]' -ItemType $_ -Message 'Access denied')

        $actual | Should -HaveCount 1
        ($actual[0].PSObject.Properties.Name -join ',') |
        Should -Be 'DateTime,ComputerName,Type,FullName,CreationTime,Action,Error'
        $actual[0].DateTime | Should -BeOfType ([datetime])
        $actual[0].ComputerName | Should -Be $env:COMPUTERNAME
        $actual[0].Type | Should -Be $_
        $actual[0].FullName | Should -Be 'z:\folder\[item]'
        $actual[0].CreationTime | Should -BeNullOrEmpty
        $actual[0].Action | Should -BeNullOrEmpty
        $actual[0].Error | Should -BeOfType ([string])
        $actual[0].Error | Should -Be 'Access denied'
    }

    It 'allows an enumeration error without a target path and emits no warning itself' {
        $actual = New-ReadErrorResultHC -FullName '' -ItemType FilesInFolder -Message 'Read failed' 3>&1

        @($actual) | Should -HaveCount 1
        $actual.FullName | Should -Be ''
        $actual.Error | Should -Be 'Read failed'
    }
}
Describe 'calendar cutoff boundaries' {
    It 'preserves <Unit> boundaries for <Today> based on <BasedOn>' -ForEach @(
        @{ Unit = 'Day'; Today = '2026-10-02T15:30:00'; Cutoff = '2026-10-02'; BasedOn = 'CreationTime' }
        @{ Unit = 'Day'; Today = '2026-10-02T15:30:00'; Cutoff = '2026-10-02'; BasedOn = 'LastWriteTime' }
        @{ Unit = 'Month'; Today = '2024-03-31T15:30:00'; Cutoff = '2024-03-01'; BasedOn = 'CreationTime' }
        @{ Unit = 'Month'; Today = '2024-03-31T15:30:00'; Cutoff = '2024-03-01'; BasedOn = 'LastWriteTime' }
        @{ Unit = 'Month'; Today = '2026-01-01T15:30:00'; Cutoff = '2026-01-01'; BasedOn = 'LastWriteTime' }
        @{ Unit = 'Year'; Today = '2026-10-02T15:30:00'; Cutoff = '2026-01-01'; BasedOn = 'CreationTime' }
        @{ Unit = 'Year'; Today = '2026-10-02T15:30:00'; Cutoff = '2026-01-01'; BasedOn = 'LastWriteTime' }
    ) {
        $testRoot = (New-Item "TestDrive:/cutoff_$([guid]::NewGuid())" -ItemType Directory).FullName
        $testBefore = New-Item "$testRoot/before.txt" -ItemType File
        $testAt = New-Item "$testRoot/at.txt" -ItemType File
        $testAfter = New-Item "$testRoot/after.txt" -ItemType File
        $testBefore.$BasedOn = ([datetime]$Cutoff).AddTicks(-1)
        $testAt.$BasedOn = [datetime]$Cutoff
        $testAfter.$BasedOn = ([datetime]$Cutoff).AddTicks(1)
        Mock Get-Date { [datetime]$Today }

        $actual = @(. $testScript -Type FilesInFolder -Path $testRoot -OlderThanUnit $Unit -OlderThanQuantity 1 -OlderThanBasedOn $BasedOn)

        $actual | Should -HaveCount 1
        $actual[0].FullName | Should -Be $testBefore.FullName
        $testBefore.FullName | Should -Not -Exist
        $testAt.FullName | Should -Exist
        $testAfter.FullName | Should -Exist
    }

    It 'rejects an overflowing <_> cutoff before deleting anything' -ForEach @('Day', 'Month', 'Year') {
        $testFile = New-Item "TestDrive:/overflow_$_.txt" -ItemType File

        { . $testScript -Type File -Path $testFile.FullName -OlderThanUnit $_ -OlderThanQuantity ([int]::MaxValue) -OlderThanBasedOn CreationTime } |
        Should -Throw '*Invalid retention period*'

        $testFile.FullName | Should -Exist
    }
}
Describe 'disappearing empty subfolders' {
    It 'skips an enumeration error only after confirming the subfolder is missing' {
        $testRoot = (New-Item 'TestDrive:/disappeared-enumeration' -ItemType Directory).FullName
        $testGone = New-Item "$testRoot/Gone" -ItemType Directory
        $testOther = New-Item "$testRoot/Other" -ItemType Directory
        Mock Get-ChildItem {
            $testGone.Delete()
            Write-Error 'Folder disappeared during enumeration' -Category ObjectNotFound -TargetObject $testGone.FullName
            $testOther
        }
        $testWarnings = @()

        $actual = @(& $testScript -Type EmptyFolders -Path $testRoot -WarningVariable testWarnings -WarningAction SilentlyContinue 2>&1)

        $actual | Should -HaveCount 1
        $actual[0].FullName | Should -Be $testOther.FullName
        $actual[0].Action | Should -Be 'Removed'
        $testWarnings | Should -BeNullOrEmpty
    }
    It 'still reports a missing configured root' {
        $testRoot = Join-Path $TestDrive 'missing-root'
        $testWarnings = @()

        $actual = @(& $testScript -Type EmptyFolders -Path $testRoot -WarningVariable testWarnings -WarningAction SilentlyContinue)

        $actual | Should -HaveCount 1
        $actual[0].FullName | Should -Be $testRoot
        $actual[0].Error | Should -Not -BeNullOrEmpty
        $testWarnings | Should -Not -BeNullOrEmpty
    }
    It 'keeps the error when the configured root also disappears' {
        $testRoot = (New-Item 'TestDrive:/disappeared-root' -ItemType Directory).FullName
        $testGone = New-Item "$testRoot/Gone" -ItemType Directory
        $testScriptText = (Get-Content -LiteralPath $testScript -Raw).Replace(
            '$iterator = $Folder.EnumerateFileSystemInfos().GetEnumerator()',
            '$Folder.Delete(); $Folder.Parent.Delete(); $iterator = $Folder.EnumerateFileSystemInfos().GetEnumerator()'
        )

        $actual = @(& ([scriptblock]::Create($testScriptText)) -Type EmptyFolders -Path $testRoot -WarningAction SilentlyContinue)

        $testRoot | Should -Not -Exist
        $actual | Should -HaveCount 1
        $actual[0].FullName | Should -Be $testGone.FullName
        $actual[0].Error | Should -Not -BeNullOrEmpty
    }
    It 'skips a folder that disappears before <Stage> and continues cleanup' -ForEach @(
        @{ Stage = 'inspection'; Anchor = '$iterator = $Folder.EnumerateFileSystemInfos().GetEnumerator()'; Injection = 'if ($Folder.Name -eq ''Gone'') { $Folder.Delete() }; ' }
        @{ Stage = 'deletion'; Anchor = '$emptyFolder.Delete()'; Injection = 'if ($emptyFolder.Name -eq ''Gone'') { $emptyFolder.Delete() }; ' }
    ) {
        $testRoot = (New-Item "TestDrive:/disappeared-$Stage" -ItemType Directory).FullName
        $testGone = New-Item "$testRoot/Gone" -ItemType Directory
        $testOther = New-Item "$testRoot/Other" -ItemType Directory
        $testScriptText = (Get-Content -LiteralPath $testScript -Raw).Replace($Anchor, "$Injection$Anchor")
        $testScriptText | Should -Not -Be (Get-Content -LiteralPath $testScript -Raw)
        $testWarnings = @()

        $actual = @(& ([scriptblock]::Create($testScriptText)) -Type EmptyFolders -Path $testRoot -WarningVariable testWarnings -WarningAction SilentlyContinue 2>&1)

        $testRoot | Should -Exist
        $testGone.FullName | Should -Not -Exist
        $testOther.FullName | Should -Not -Exist
        $actual | Should -HaveCount 1
        $actual[0].FullName | Should -Be $testOther.FullName
        $actual[0].Action | Should -Be 'Removed'
        $testWarnings | Should -BeNullOrEmpty
        @($actual | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] }) | Should -HaveCount 0
    }
}
Describe 'lazy emptiness checks' {
    It 'checks one entry and disposes the iterator, including failure: <_>' -ForEach @($false, $true) {
        $testRoot = (New-Item "TestDrive:/lazy_$_" -ItemType Directory).FullName
        $testFolder = New-Item "$testRoot/child" -ItemType Directory
        $testIterator = [pscustomobject]@{ Calls = 0; Disposed = $false; Fail = $_ }
        $testIterator | Add-Member ScriptMethod MoveNext {
            $this.Calls++
            if ($this.Fail -or ($this.Calls -gt 1)) { throw 'Enumeration failed' }
            $true
        }
        $testIterator | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
        $testEnumerable = [pscustomobject]@{ Iterator = $testIterator }
        $testEnumerable | Add-Member ScriptMethod GetEnumerator { $this.Iterator }
        $testFolder | Add-Member NoteProperty TestEnumerable $testEnumerable
        $testFolder | Add-Member ScriptMethod EnumerateFileSystemInfos { $this.TestEnumerable } -Force
        Mock Get-ChildItem { $testFolder }

        $actual = @(. $testScript -Type EmptyFolders -Path $testRoot -WarningAction SilentlyContinue)

        $testIterator.Calls | Should -Be 1
        $testIterator.Disposed | Should -BeTrue
        $testFolder.FullName | Should -Exist
        if ($testIterator.Fail) {
            $actual | Should -HaveCount 1
            $actual[0].Error | Should -BeLike '*Enumeration failed*'
        }
        else { $actual | Should -HaveCount 0 }
    }
}
Describe 'excluded subtree traversal' {
    It 'does not enumerate excluded subtrees for <_>' -ForEach @('FilesInFolder', 'EmptyFolders') {
        $testType = $_
        $testRoot = (New-Item "TestDrive:/pruned_$testType" -ItemType Directory).FullName
        $testKeep = New-Item "$testRoot/Keep/deep/protected.txt" -ItemType File -Force
        $testEmpty = New-Item "$testRoot/Other/empty" -ItemType Directory -Force
        $testRemove = New-Item "$testRoot/Other/remove.txt" -ItemType File
        $testGetChildItem = Get-Command Get-ChildItem -CommandType Cmdlet
        Mock Get-ChildItem {
            $testEnumeration = @{ LiteralPath = $LiteralPath; Force = $true }
            if ($PesterBoundParameters['File']) { $testEnumeration.File = $true }
            if ($PesterBoundParameters['Directory']) { $testEnumeration.Directory = $true }
            if ($PesterBoundParameters['Recurse']) { $testEnumeration.Recurse = $true }
            & $testGetChildItem @testEnumeration
        }

        . $testScript -Type $testType -Path $testRoot -ExcludeFolder "$testRoot/Keep" -Recurse $true -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime

        $testKeep.FullName | Should -Exist
        if ($testType -eq 'FilesInFolder') { $testRemove.FullName | Should -Not -Exist }
        else { $testEmpty.FullName | Should -Not -Exist }
        Should -Invoke Get-ChildItem -Times 0 -Exactly -Scope It -ParameterFilter { $PesterBoundParameters['Recurse'] }
        Should -Invoke Get-ChildItem -Times 0 -Exactly -Scope It -ParameterFilter {
            $LiteralPath -like "$testRoot\Keep*"
        }
        Should -Invoke Get-ChildItem -Times 1 -Exactly -Scope It -ParameterFilter {
            $LiteralPath -eq "$testRoot\Other"
        }
        if ($testType -eq 'EmptyFolders') {
            Should -Invoke Get-ChildItem -Times 0 -Exactly -Scope It -ParameterFilter {
                -not $PesterBoundParameters['Directory']
            }
        }
        else {
            Should -Invoke Get-ChildItem -Times 0 -Exactly -Scope It -ParameterFilter {
                $PesterBoundParameters.ContainsKey('Directory')
            }
        }
    }

    It 'does not follow a junction outside the tree' {
        $testRoot = (New-Item 'TestDrive:/junction-root' -ItemType Directory).FullName
        $testOutside = New-Item 'TestDrive:/junction-outside/protected.txt' -ItemType File -Force
        $testLink = New-Item "$testRoot/link" -ItemType Junction -Value $testOutside.Directory.FullName
        try {
            . $testScript -Type FilesInFolder -Path $testRoot -ExcludeFolder "$testRoot/unused" -Recurse $true -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime

            $testOutside.FullName | Should -Exist
        }
        finally { $testLink.Delete() }
    }

    It 'reports an unreadable included sibling for <_>' -ForEach @('FilesInFolder', 'EmptyFolders') {
        $testType = $_
        $testRoot = (New-Item "TestDrive:/pruned-denied_$testType" -ItemType Directory).FullName
        $testDenied = (New-Item "$testRoot/denied" -ItemType Directory).FullName
        $testUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        $testRule = [System.Security.AccessControl.FileSystemAccessRule]::new($testUser, 'ListDirectory', 'Deny')
        $testAcl = Get-Acl -LiteralPath $testDenied
        $testAcl.AddAccessRule($testRule)
        Set-Acl -LiteralPath $testDenied -AclObject $testAcl
        try {
            $actual = @(. $testScript -Type $testType -Path $testRoot -ExcludeFolder "$testRoot/unused" -Recurse $true -OlderThanUnit Day -OlderThanQuantity 0 -OlderThanBasedOn CreationTime -WarningAction SilentlyContinue)
        }
        finally {
            $testAcl.RemoveAccessRule($testRule) | Out-Null
            Set-Acl -LiteralPath $testDenied -AclObject $testAcl
        }

        @($actual | Where-Object FullName -EQ $testDenied) | Should -HaveCount 1
        ($actual | Where-Object FullName -EQ $testDenied).Error | Should -Not -BeNullOrEmpty
        $testDenied | Should -Exist
    }
}
Describe 'single-pass empty-folder cleanup' {
    It 'enumerates a deep tree once and removes every child before its parent' {
        $testRoot = (New-Item 'TestDrive:/deep-tree' -ItemType Directory).FullName
        $testNested = $testRoot
        foreach ($level in 1..40) { $testNested = Join-Path $testNested 'd' }
        $null = New-Item $testNested -ItemType Directory -Force
        $testGetChildItem = Get-Command Get-ChildItem -CommandType Cmdlet
        Mock Get-ChildItem { & $testGetChildItem -LiteralPath $LiteralPath -Directory -Recurse -Force }

        $actual = @(. $testScript -Type EmptyFolders -Path $testRoot)

        $actual | Should -HaveCount 40
        @($actual.FullName | Select-Object -Unique) | Should -HaveCount 40
        $actual[0].FullName | Should -Be $testNested
        $actual[-1].FullName | Should -Be (Join-Path $testRoot 'd')
        $actual.Error | Where-Object { $_ } | Should -BeNullOrEmpty
        $testRoot | Should -Exist
        Should -Invoke Get-ChildItem -Times 1 -Exactly -Scope It
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
