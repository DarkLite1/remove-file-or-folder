#Requires -Version 7

<#
.SYNOPSIS
    Convert an input file from the old 'Remove' format to the 'Tasks' format.

.DESCRIPTION
    Paths with the same computer and settings are grouped in one task. A
    folder in 'Remove.EmptyFolders' that is also in 'Remove.FilesInFolder'
    becomes 'RemoveEmptyFolders: true' on that folder; the other empty folders
    get a task without 'OlderThan'.

    Sets OlderThan.BasedOn to CreationTime to preserve the old date filter.

    'Settings' and 'PSSessionConfiguration' are copied from the template file.
    Only 'Settings.ScriptName' (the file name) and 'Settings.SendMail.To' and
    'When' (from the old file) are filled in, so check the other settings.

    'MaxConcurrentJobs' becomes both 'JobsTotal' and 'JobsPerComputer', which
    keeps the old behavior. If omitted or zero, both limits become 1.

    Writes the converted configuration only; it does not run removal tasks.
    Review the settings and paths before using the result with Main.ps1.

.PARAMETER Path
    The input file in the old format.

.PARAMETER Destination
    Output JSON file. An existing file is overwritten. Its parent folder
    must already exist.

.PARAMETER TemplateFile
    File to copy 'Settings' and 'PSSessionConfiguration' from. Defaults to
    Example.json in the repository root. Mail recipients and send policy
    remain from this template when the old file does not supply them.

.EXAMPLE
    & '.\Tools\Convert-InputFile.ps1' -Path 'C:\old\BNL CL.json' -Destination 'C:\new\BNL CL.json'
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [String]$Destination,
    [String]$TemplateFile = (Join-Path (Split-Path $PSScriptRoot) 'Example.json')
)

$ErrorActionPreference = 'Stop'

$old = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
$template = Get-Content -LiteralPath $TemplateFile -Raw -Encoding UTF8 | ConvertFrom-Json

if (-not $old.Remove) {
    throw "File '$Path' has no 'Remove' property, it is not in the old format"
}

function Get-KeyHC {
    <#
    .SYNOPSIS
        Build a case-insensitive lookup key from a computer name and path.
        Path separators and trailing slashes are not normalized.
    #>
    param ([String]$ComputerName, [String]$Path)
    '{0}|{1}' -f "$ComputerName".ToLower(), $Path.ToLower()
}

function ConvertTo-PathEntryHC {
    <#
    .SYNOPSIS
        Return a Name/Path object when a friendly name is supplied;
        otherwise return the path string.
    #>
    param ([String]$Name, [String]$Path)

    if ($Name) {
        [ordered]@{ Name = $Name; Path = $Path }
    }
    else {
        $Path
    }
}

#region Empty folders
$emptyFolders = [ordered]@{}

foreach ($item in $old.Remove.EmptyFolders) {
    $emptyFolders[(Get-KeyHC $item.ComputerName $item.Path)] = $item
}

$emptyFoldersUsed = @{}
#endregion

#region Group paths with the same settings
$groups = [ordered]@{}

function Add-PathHC {
    <#
    .SYNOPSIS
        Add a path entry to the shared conversion group for its settings.

    .DESCRIPTION
        Uses the converter's groups dictionary. Creates the group from Task
        when needed, then adds PathEntry to the Files or Folders list named
        by Task.ListName.
    #>
    param ([String]$GroupKey, [hashtable]$Task, $PathEntry)

    if (-not $groups.Contains($GroupKey)) {
        $groups[$GroupKey] = $Task
    }
    $groups[$GroupKey][$Task.ListName] += , $PathEntry
}

foreach ($item in $old.Remove.File) {
    $groupKey = 'File|{0}|{1}|{2}' -f
    $item.ComputerName, $item.OlderThan.Quantity, $item.OlderThan.Unit

    $params = @{
        GroupKey  = $groupKey
        Task      = @{
            ListName     = 'Files'
            ComputerName = $item.ComputerName
            OlderThan    = $item.OlderThan
            Files        = @()
        }
        PathEntry = ConvertTo-PathEntryHC $item.Name $item.Path
    }
    Add-PathHC @params
}

foreach ($item in $old.Remove.FilesInFolder) {
    $key = Get-KeyHC $item.ComputerName $item.Path
    $emptyFolder = $emptyFolders[$key]
    $removeEmptyFolders = $null -ne $emptyFolder

    if ($removeEmptyFolders) {
        $emptyFoldersUsed[$key] = $true
    }

    $name = if ($item.Name) { $item.Name } else { $emptyFolder.Name }

    $groupKey = 'Folder|{0}|{1}|{2}|{3}|{4}' -f
    $item.ComputerName, $item.OlderThan.Quantity, $item.OlderThan.Unit,
    $item.Recurse, $removeEmptyFolders

    $params = @{
        GroupKey  = $groupKey
        Task      = @{
            ListName           = 'Folders'
            ComputerName       = $item.ComputerName
            OlderThan          = $item.OlderThan
            Recurse            = [bool]$item.Recurse
            RemoveEmptyFolders = $removeEmptyFolders
            Folders            = @()
        }
        PathEntry = ConvertTo-PathEntryHC $name $item.Path
    }
    Add-PathHC @params
}

foreach ($entry in $emptyFolders.GetEnumerator()) {
    if ($emptyFoldersUsed[$entry.Key]) { continue }

    $item = $entry.Value

    $params = @{
        GroupKey  = 'EmptyFolders|{0}' -f $item.ComputerName
        Task      = @{
            ListName           = 'Folders'
            ComputerName       = $item.ComputerName
            RemoveEmptyFolders = $true
            Folders            = @()
        }
        PathEntry = ConvertTo-PathEntryHC $item.Name $item.Path
    }
    Add-PathHC @params
}
#endregion

#region Create tasks
$tasks = foreach ($group in $groups.Values) {
    $task = [ordered]@{ ComputerName = $group.ComputerName }

    $task[$group.ListName] = $group[$group.ListName]

    if ($group.OlderThan) {
        # the old format always compared the creation time
        $task.OlderThan = [ordered]@{
            Quantity = $group.OlderThan.Quantity
            Unit     = $group.OlderThan.Unit
            BasedOn  = 'CreationTime'
        }
    }
    if ($group.ListName -eq 'Folders') {
        if ($group.OlderThan) {
            $task.Recurse = $group.Recurse
        }
        $task.RemoveEmptyFolders = $group.RemoveEmptyFolders
    }

    $task
}
#endregion

#region Create the new file
$settings = $template.Settings
$settings.ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($Path)

if ($old.SendMail.To) {
    $settings.SendMail.To = @($old.SendMail.To)
}
if ($old.SendMail.When) {
    $settings.SendMail.When = switch ($old.SendMail.When) {
        'OnlyOnError' { 'OnError' }
        'OnlyOnErrorOrAction' { 'OnErrorOrAction' }
        default { $_ }
    }
}

$maxConcurrentJobs = if ($old.MaxConcurrentJobs) {
    [int]$old.MaxConcurrentJobs
}
else {
    1
}

$new = [ordered]@{
    MaxConcurrent          = [ordered]@{
        JobsTotal       = $maxConcurrentJobs
        JobsPerComputer = $maxConcurrentJobs
    }
    Tasks                  = @($tasks)
    PSSessionConfiguration = $template.PSSessionConfiguration
    Settings               = $settings
}

$new | ConvertTo-Json -Depth 10 |
Out-File -LiteralPath $Destination -Encoding utf8
#endregion
