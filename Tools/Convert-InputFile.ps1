#Requires -Version 7

<#
.SYNOPSIS
    Convert an input file from the old 'Remove' format to the 'Tasks' format.

.DESCRIPTION
    Paths with the same computer and settings are grouped in one task. A
    folder in 'Remove.EmptyFolders' that is also in 'Remove.FilesInFolder'
    becomes 'RemoveEmptyFolders: true' on that folder; the other empty folders
    get a task without 'OlderThan'.

    'Settings' and 'PSSessionConfiguration' are copied from the template file.
    Only 'Settings.ScriptName' (the file name) and 'Settings.SendMail.To' and
    'When' (from the old file) are filled in, so check the other settings.

    'MaxConcurrentJobs' becomes both 'JobsTotal' and 'JobsPerComputer', which
    keeps the old behavior.

.PARAMETER Path
    The input file in the old format.

.PARAMETER Destination
    The file to create in the new format.

.PARAMETER TemplateFile
    The file to copy 'Settings' and 'PSSessionConfiguration' from.

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
    param ([String]$ComputerName, [String]$Path)
    '{0}|{1}' -f "$ComputerName".ToLower(), $Path.ToLower()
}

function ConvertTo-PathEntryHC {
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
        $task.OlderThan = [ordered]@{
            Quantity = $group.OlderThan.Quantity
            Unit     = $group.OlderThan.Unit
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
