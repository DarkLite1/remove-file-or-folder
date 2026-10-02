# 🧹 PowerShell Remove File or Folder

> **Keep shares, log folders and drop zones clean, on one machine or hundreds.**

This PowerShell script removes old files, the content of folders and empty folders on local or remote machines, driven by a single JSON configuration file. It runs jobs in parallel with per-computer limits, retries dropped remote connections, logs everything and sends a clear summary by e-mail.

---

## ✨ Why use this?

- 🗑️ **Flexible removal:** Remove specific files, the files in folders (optionally recursive), the empty folders below a path, or files and empty folders in one go.
- 📋 **Compact configuration:** One task holds a list of paths that share the same computer and settings.
- ⏳ **Age based:** Only remove files older than a number of days, months or years, based on their creation date. Use `0` to remove everything.
- ⚡ **Parallel execution:** Limit the total number of jobs (`JobsTotal`) and the number of jobs per computer (`JobsPerComputer`), so busy servers are never overloaded.
- 🔁 **Resilient:** Remote jobs that fail because of a dropped WinRM connection are retried automatically.
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
      "OlderThan": { "Quantity": 30, "Unit": "Day" },
      "Recurse": true,
      "RemoveEmptyFolders": true
    },
    {
      "ComputerName": null,
      "Files": [{ "Name": "Upload log", "Path": "\\\\contoso\\sftp\\upload.log" }],
      "OlderThan": { "Quantity": 0, "Unit": "Day" }
    }
  ],
  "Settings": {
    "ScriptName": "Remove old logs",
    "SendMail": {
      "When": "OnError",
      "To": ["admin@example.com"]
    }
  }
}
```

The mail server, log folder and event log settings are left out above for brevity; they are required, see `Example.json`.

### 📋 Tasks

Every task targets one computer and applies the same settings to a list of paths. Use as many tasks as needed, for example one per computer and retention period.

| Property             | Description                                                                                                                              |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| `ComputerName`       | The computer that executes the removal. Required for local paths like `D:\Logs`, use `null` for UNC paths and `localhost` for this computer. |
| `Files`              | The files to remove. Requires `OlderThan`.                                                                                                |
| `Folders`            | The folders to clean up. Requires `RemoveEmptyFolders`, and `Recurse` when `OlderThan` is used.                                          |
| `ExcludeFolders`     | Optional subfolders of `Folders` to skip. Nothing inside them is removed, and they are never removed as empty folders.                   |
| `OlderThan`          | Remove the files older than this. Leave it out for `Folders` to only remove empty folders.                                               |
| `Recurse`            | `true` also removes the files in the subfolders.                                                                                         |
| `RemoveEmptyFolders` | `true` removes the empty folders below each folder, after all files are removed. The folder itself is never removed.                     |

A path in `Files` or `Folders` is a plain string, or an object `{ "Name": "...", "Path": "..." }` when the e-mail should show a friendly name instead of the path.

Every path is processed as a separate job, so `MaxConcurrent` applies per path. Empty folders are always removed after all file removals have finished.

### 🔄 Converting input files from the old format

Input files with a `Remove` section (`File`, `FilesInFolder`, `EmptyFolders`) are no longer supported. Convert them with:

```powershell
& '.\Tools\Convert-InputFile.ps1' -Path 'C:\old\BNL CL.json' -Destination 'C:\new\BNL CL.json'
```

The converter groups the paths with the same computer and settings in one task, and turns an `EmptyFolders` entry for a folder that is also in `FilesInFolder` into `RemoveEmptyFolders: true`. `Settings` are copied from `Example.json` with the old `SendMail.To` and `SendMail.When`, so check the mail server, log folder and event log settings afterwards.

### ⏳ How `OlderThan` works

Files are selected on their **creation date**. `Quantity` `0` removes all files, regardless of their age.

`Month` and `Year` compare **calendar periods** by design, not an exact number of days. Only the month or year of the creation date counts, the day is ignored:

| `OlderThan`         | Run on 1 October 2026 removes files created |
| ------------------- | ------------------------------------------- |
| `1 Day`             | on or before 30 September 2026              |
| `30 Day`            | on or before 1 September 2026               |
| `1 Month`           | in September 2026 or earlier                |
| `1 Year`            | in 2025 or earlier                          |

So `1 Month` removes a file created on 30 September, even though it is only one day old. When you need an exact age, like 30 days, use `Day` with `30` instead of `Month` with `1`.

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
