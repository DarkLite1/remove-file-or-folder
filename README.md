# 🧹 PowerShell Remove File or Folder

> **Keep shares, log folders and drop zones clean, on one machine or hundreds.**

This PowerShell script removes old files and empty subfolders on local or remote Windows machines, driven by a single JSON configuration file. It runs jobs in parallel with per-computer limits and can save logs and send a summary by e-mail.

---

## ✨ Why use this?

- 🗑️ **Flexible removal:** Remove specific files, the files in folders (optionally recursive), the empty folders below a path, or files and empty folders in one go.
- 📋 **Compact configuration:** One task holds a list of paths that share the same computer and settings.
- ⏳ **Age based:** Remove files using calendar-day, month or year cutoffs, based on their creation or last write date. Use `0` to remove all selected files, while still respecting exclusions.
- ⚡ **Parallel execution:** Limit the total number of jobs (`JobsTotal`) and the number of jobs per computer (`JobsPerComputer`), so busy servers are never overloaded.
- 🔁 **Retries:** Jobs that fail with WinRM abort error 995, or the message "I/O operation has been aborted", get up to three total attempts with five seconds between attempts. Other errors are reported without retrying.
- 🌍 **Remote ready:** Use local paths with a `ComputerName` (PowerShell remoting) or plain UNC paths.
- 📊 **Logging:** Excel overview of every removed item, a system errors log and optional Windows Event Log entries.
- 📬 **E-mail alerts:** Get a summary when items are removed, when errors occur, always or never.
- 🔐 **Secure & portable:** Use `ENV:` environment variables for sensitive values like SMTP credentials.

---

## 🚀 Getting Started

### 1. Prerequisites

- **PowerShell 7.x+**
- **ImportExcel** module:
  ```powershell
  Install-Module -Name 'ImportExcel' -Scope 'AllUsers'
  ```
- **PowerShell remoting** on remote computers, with the endpoint set in `PSSessionConfiguration` (default `PowerShell.7`).
- **Email dependencies (optional)**: install MailKit and MimeKit from an elevated PowerShell prompt:
  ```powershell
  Install-Package -Name 'MailKit','MimeKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  ```

### 2. Installation

```bash
git clone https://github.com/DarkLite1/remove-file-or-folder.git
cd remove-file-or-folder
```

### 3. Configuration

Copy `Example.json` to `MyConfig.json` and adjust it. `Example.json` contains a task for every use case, and every property is explained in its `?` section. A minimal example:

```json
{
  "MaxConcurrent": {
    "JobsTotal": 4,
    "JobsPerComputer": 2
  },
  "Tasks": [
    {
      "ComputerName": "SERVER1",
      "Folders": ["D:\\Application\\Logs", "D:\\Application\\Archive"],
      "OlderThan": { "Quantity": 30, "Unit": "Day", "BasedOn": "LastWriteTime" },
      "Recurse": true,
      "RemoveEmptyFolders": true
    },
    {
      "ComputerName": null,
      "Files": [{ "Name": "Upload log", "Path": "\\\\contoso\\sftp\\upload.log" }],
      "OlderThan": { "Quantity": 0, "Unit": "Day", "BasedOn": "CreationTime" }
    }
  ],
  "Settings": {
    "ScriptName": "Remove old logs",
    "SendMail": {
      "When": "Never"
    },
    "SaveInEventLog": {
      "Save": false
    }
  }
}
```

Replace the computer names and paths before running this example. It disables email and event logging and does not save log files. To enable reporting, use the settings in `Example.json`:

- **Email:** Set `Settings.SendMail.When` to `Always`, `OnError` or `OnErrorOrAction`. Then provide `From`, at least one `To` or `Bcc` recipient, `Smtp.ServerName`, `Smtp.Port`, and both `AssemblyPath` values. SMTP authentication is optional: leave both `UserName` and `Password` empty, or provide both. Providing only one causes an error when sending mail.
- **Log files:** Set `Settings.SaveLogFiles.Where.Folder`. Omit it or leave it empty to disable file logging. When an email is generated, its exact HTML body is saved in this folder as `<yyyy_MM_dd_HHmmss> - <ScriptName> (<configuration name>) - Mail.html`, using the same prefix as the other logs. This copy is not attached to the email and remains available if sending fails.
- **Event log:** Always include `Settings.SaveInEventLog.Save`. When it is `true`, also provide `LogName`.

Missing or inaccessible paths and item-removal failures appear in the Excel **Overview** worksheet's **Error** column. Job-execution failures appear in the **Errors** worksheet with **Stage**, **TargetObject**, **FullyQualifiedErrorId**, **ExceptionType**, **ScriptStackTrace** and **PositionMessage** diagnostics when available. **Path** identifies the configured task root; **TargetObject** identifies the object associated with the original error. Neither kind of error is duplicated in JSON. The system errors JSON log is reserved for script, configuration and reporting failures. All errors still count toward error notifications and exit code 1.

`Example.json` enables email, file logging and event logging. Its server names, addresses and assembly paths are placeholders to adjust for your environment.

The email and saved HTML show one row per entry in `Tasks`, listing its paths once and combining the removal/error counts and descriptions of its file and empty-folder jobs. Separate task entries remain separate even when they target the same path. A task spanning multiple UNC servers appears once under a combined server heading. Excel retains the detailed per-item and per-job records.

Before reporting a subfolder read or deletion failure during empty-folder cleanup, the worker rechecks the path. If the subfolder is confirmed missing and the configured root remains accessible, it is skipped without reporting an error or claiming a removal. Access failures, inconclusive checks and missing configured roots remain errors.

Identical folder-read errors for the same computer and path within one input task are counted once across Excel, email and event reporting. In Excel **Overview**, **Type** lists the affected operations, such as `FilesInFolder, EmptyFolders`; the first occurrence's timestamp and retention settings are retained. Different errors, separate input tasks, deletion failures and successful removals stay separate. Email totals include the unique item errors in **Overview** plus job failures in **Errors**. Script, configuration and reporting failures are counted additionally and listed separately in the email and system errors JSON log.

### 📋 Tasks

Every task targets one computer and applies the same settings to a list of paths. Choose either `Files` or `Folders` in each task, never both. Use as many tasks as needed, for example one per computer and retention period.

| Property             | Description                                                                                                                              |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| `ComputerName`       | The computer that executes the removal. Required for local paths like `D:\Logs`, use `null` for UNC paths and `localhost` for this computer. |
| `Files`              | The files to remove. Requires `OlderThan`.                                                                                                |
| `Folders`            | The folders to clean up. Requires `RemoveEmptyFolders`, and `Recurse` when `OlderThan` is used.                                          |
| `Exclude.Folders`    | Optional subfolders of `Folders` to skip. Nothing inside them is removed, and they are never removed as empty folders.                   |
| `Exclude.Files`      | Optional files below `Folders` that are never removed, regardless of their age. Requires `OlderThan`.                                    |
| `Exclude.Attributes` | Optional array: `[]`, `["Hidden"]`, `["System"]` or `["Hidden", "System"]`. Protects matching files and skips matching folder trees. Defaults to `[]`. |
| `OlderThan`          | Remove the files older than this, based on their `CreationTime` or `LastWriteTime`. Leave it out for `Folders` to only remove empty folders. |
| `Recurse`            | Required only for `Folders` with `OlderThan`. `true` includes files in subfolders; `false` selects files directly in the folder. Not allowed on other tasks. |
| `RemoveEmptyFolders` | Required for `Folders` only. `true` removes empty subfolders at every depth after the file-removal phase, independently of `Recurse`. The root folder is never removed. |

A path in `Files` or `Folders` is a plain string, or an object `{ "Name": "...", "Path": "..." }` when the e-mail should show a friendly name above the path. Both remain visible.

Every path is processed as a separate job, so `MaxConcurrent` applies per path. Empty folders are always removed after all file removals have finished.

Group exclusions in an optional `Exclude` object on the relevant task. For a task targeting `D:\Application\Logs`:

```json
"Exclude": {
  "Attributes": ["Hidden", "System"],
  "Folders": ["D:\\Application\\Logs\\Keep"],
  "Files": ["D:\\Application\\Logs\\state.json"]
}
```

Each property is optional and takes an array. Omit unused properties or use `[]`; omitting `Exclude` or using `{}` adds no exclusions. `Exclude.Folders` and `Exclude.Files` only apply to tasks targeting folders; `Exclude.Files` also requires `OlderThan`. Paths must be below a configured task folder. Unknown properties and invalid values are rejected before cleanup starts.

Matching **either** attribute is enough. Matching files are not deleted; matching folders and all their contents are skipped before traversal in both cleanup phases. This also skips a configured root folder when its attributes match. Explicit `Files` entries are checked against the file's own attributes, not its ancestors. Retained hidden/system files still make a folder nonempty, so its parent is not removed by empty-folder cleanup.

Omitting `Exclude.Attributes` or using `[]` preserves the existing behavior, including hidden and system items. Use `["System"]` to protect system items while still cleaning hidden application logs. Exclusions are shown in the email description; skipped items do not count as removals or errors. Access errors on other paths, or failures to read a path's attributes, remain visible. Attribute exclusions work alongside `Exclude.Files` and `Exclude.Folders` and still apply when `OlderThan.Quantity` is `0`.

### 🔄 Converting input files from the old format

Input files with a `Remove` section (`File`, `FilesInFolder`, `EmptyFolders`) are no longer supported. Convert them with:

```powershell
& '.\Tools\Convert-InputFile.ps1' -Path 'C:\old\BNL CL.json' -Destination 'C:\new\BNL CL.json'
```

The converter groups the paths with the same computer and settings in one task, and turns an `EmptyFolders` entry for a folder that is also in `FilesInFolder` into `RemoveEmptyFolders: true`. `OlderThan.BasedOn` is set to `CreationTime`, which is what the old format used. `Settings` are copied from `Example.json` with the old `SendMail.To` and `SendMail.When`, so check the mail server, log folder and event log settings afterwards.

### ⏳ How `OlderThan` works

`OlderThan.BasedOn` decides which file date is compared, it is required:

| `BasedOn`       | Meaning                                                    | Use it for                                                                                     |
| --------------- | ---------------------------------------------------------- | ---------------------------------------------------------------------------------------------- |
| `CreationTime`  | When the file was created on this disk                     | Files that are written once, like scans, exports or drop files                                 |
| `LastWriteTime` | When the content was last changed                          | Files that are updated over time, like logs or history files: they are kept while still in use |

A copied file gets a new `CreationTime` but keeps its `LastWriteTime`. So with `LastWriteTime` a file that was just copied can already be old enough to be removed.

`Quantity` must be a whole number of `0` or greater. `0` disables the age filter: all files selected by the task are eligible, but exclusions and `Recurse` still apply. Age settings never select folders for deletion; folders must be empty.

All units compare **calendar periods**, not elapsed time. `Day` includes the entire cutoff day, ignoring the time of day. `Month` includes the entire cutoff month, and `Year` includes the entire cutoff year:

| `OlderThan`         | Run on 1 October 2026 removes files dated |
| ------------------- | ----------------------------------------- |
| `1 Day`             | on or before 30 September 2026            |
| `30 Day`            | on or before 1 September 2026             |
| `1 Month`           | in September 2026 or earlier              |
| `1 Year`            | in 2025 or earlier                        |

For example, `1 Day` run at 00:05 can remove a file last written at 23:55 the previous day, only ten minutes earlier. Likewise, `1 Month` on 1 October can remove a file dated 30 September. Use `30 Day` for a cutoff 30 calendar days back instead of `1 Month`; neither setting guarantees an exact elapsed age. These cutoffs use the clock and local timestamps of the computer executing the removal.

### Email summary

The email always includes summary counts and task results. `Settings.SendMail.Body` adds optional HTML below the title; leaving it empty does not remove the summary. The subject starts with `N removed`, adds the error count only when there are errors, and appends any text in `Settings.SendMail.Subject`.

## 💻 Usage

```powershell
& '.\Main.ps1' -ConfigurationJsonFile '.\MyConfig.json'
```

## ⏱️ Automating with Task Scheduler

- Program: `pwsh.exe`
- Arguments:

```
-Command "& 'C:\remove-file-or-folder\Main.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit $LASTEXITCODE"
```

**_Pro Tip:_** Run the task with an account that has delete permissions on all paths and remoting access to all computers.

## 🚦 Exit Codes

| Code | Status  | Description                                                                                                                                         |
| ---- | ------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0    | Success | All jobs completed without errors.                                                                                                                  |
| 1    | Error   | An invalid configuration file, a job that failed to run, an item that could not be removed (e.g. file in use) or a script failure like an e-mail error. |

## 🧪 Tests

The Pester tests are in the `Tests` folder:

```powershell
Invoke-Pester -Path '.\Tests' -Output Detailed
```

## Performance and safety

- Empty folders are enumerated once and processed deepest first. The root is preserved, and nonrecursive deletion refuses folders that gain content during cleanup.
- Excluded folder trees are skipped before descent. Paths are normalized before comparison, and file exclusions use case-insensitive exact matching.
- File candidates are streamed instead of collecting the entire tree. Eligible files have their metadata refreshed before deletion, including another age check. This narrows, but cannot eliminate, races with concurrent writers.
- Hidden and read-only items are included. Directory junctions and symbolic links encountered during traversal are not followed into other trees.
- The worker still holds directory candidates for bottom-up sorting. The main script retains removal/error results for reporting, so total job memory is not constant.

Run the repeatable local benchmark from the repository root:

```powershell
& '.\Tools\Measure-RemovalPerformance.ps1' -FileCount 10000
```

It compares the working worker with commit `3712337` (before the performance changes), uses only newly created temporary fixtures, and removes them afterward. Each scenario has an unreported warmup followed by three measured runs. Fixture preparation and verification are outside the measured interval. Every run verifies removal counts and retained file counts.

Scenarios cover deep trees, excluded trees, file filtering, wide directories, and bulk file deletion. `AllocatedMB` is total managed allocation during a run, not peak memory. Results depend on filesystem caching, storage, antivirus, and network latency; local results are not an SMB performance guarantee.

```powershell
& '.\Tools\Measure-RemovalPerformance.ps1' -Scenario FileFiltering -FileCount 100000 -Iterations 1
```
