#Requires -Version 7
#Requires -Modules Pester

BeforeAll {
    $testScript = Join-Path (Split-Path $PSScriptRoot) 'Tools\Convert-InputFile.ps1'

    $testOldFile = @{
        MaxConcurrentJobs = 3
        SendMail          = @{
            To   = @('bob@example.com')
            When = 'OnlyOnErrorOrAction'
        }
        Remove            = @{
            File          = @(
                @{ Name = 'Log'; ComputerName = $null; Path = '\\s\a.txt'; OlderThan = @{ Quantity = 0; Unit = 'Day' } }
                @{ Name = $null; ComputerName = $null; Path = '\\s\b.txt'; OlderThan = @{ Quantity = 0; Unit = 'Day' } }
            )
            FilesInFolder = @(
                @{ Name = $null; ComputerName = 'PC1'; Path = 'D:\a'; Recurse = $true; OlderThan = @{ Quantity = 3; Unit = 'Month' } }
                @{ Name = $null; ComputerName = 'PC1'; Path = 'D:\b'; Recurse = $true; OlderThan = @{ Quantity = 3; Unit = 'Month' } }
                @{ Name = $null; ComputerName = 'PC1'; Path = 'D:\c'; Recurse = $true; OlderThan = @{ Quantity = 7; Unit = 'Day' } }
            )
            EmptyFolders  = @(
                @{ Name = 'Folder A'; ComputerName = 'PC1'; Path = 'D:\a' }
                @{ Name = 'Folder B'; ComputerName = 'PC1'; Path = 'D:\b' }
                @{ Name = $null; ComputerName = $null; Path = '\\s\empty' }
            )
        }
    }

    $testOldPath = (New-Item 'TestDrive:/My input.json' -ItemType File).FullName
    $testNewPath = Join-Path (Split-Path $testOldPath) 'new.json'

    $testOldFile | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $testOldPath

    & $testScript -Path $testOldPath -Destination $testNewPath

    $actual = Get-Content -LiteralPath $testNewPath -Raw | ConvertFrom-Json
}
Describe 'Convert-InputFile' {
    It 'groups files with the same settings in one task' {
        $task = $actual.Tasks | Where-Object { $_.Files }

        @($task) | Should -HaveCount 1
        $task.Files[0].Name | Should -Be 'Log'
        $task.Files[0].Path | Should -Be '\\s\a.txt'
        $task.Files[1] | Should -Be '\\s\b.txt'
        $task.OlderThan.Quantity | Should -Be 0
    }
    It 'groups folders with the same computer and settings in one task' {
        $task = $actual.Tasks | Where-Object { $_.OlderThan.Unit -eq 'Month' }

        @($task) | Should -HaveCount 1
        $task.ComputerName | Should -Be 'PC1'
        $task.Recurse | Should -BeTrue
        @($task.Folders) | Should -HaveCount 2
    }
    It 'sets RemoveEmptyFolders for folders that were also in EmptyFolders' {
        $task = $actual.Tasks | Where-Object { $_.OlderThan.Unit -eq 'Month' }

        $task.RemoveEmptyFolders | Should -BeTrue
        $task.Folders.Name | Should -Be @('Folder A', 'Folder B')
    }
    It 'does not set RemoveEmptyFolders for the other folders' {
        $task = $actual.Tasks | Where-Object { $_.OlderThan.Quantity -eq 7 }

        $task.RemoveEmptyFolders | Should -BeFalse
        $task.Folders | Should -Be 'D:\c'
    }
    It 'creates a task without OlderThan for the remaining empty folders' {
        $task = $actual.Tasks | Where-Object { -not $_.OlderThan }

        $task.Folders | Should -Be '\\s\empty'
        $task.RemoveEmptyFolders | Should -BeTrue
        $task.PSObject.Properties.Name | Should -Not -Contain 'Recurse'
    }
    It 'converts MaxConcurrentJobs and SendMail' {
        $actual.MaxConcurrent.JobsTotal | Should -Be 3
        $actual.MaxConcurrent.JobsPerComputer | Should -Be 3
        $actual.Settings.SendMail.To | Should -Be 'bob@example.com'
        $actual.Settings.SendMail.When | Should -Be 'OnErrorOrAction'
        $actual.Settings.ScriptName | Should -Be 'My input'
    }
    It 'throws for a file that is not in the old format' {
        { & $testScript -Path $testNewPath -Destination 'TestDrive:/x.json' } |
        Should -Throw "*has no 'Remove' property*"
    }
}
