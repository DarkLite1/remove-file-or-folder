# 🧹 PowerShell Remove File or Folder

> **Keep shares, log folders and drop zones clean, on one machine or hundreds.**

This PowerShell script removes old files, the content of folders and empty folders on local or remote machines, driven by a single JSON configuration file. It runs jobs in parallel with per-computer limits, retries dropped remote connections, logs everything and sends a clear summary by e-mail.

---

## ✨ Why use this?

- 🗑️ **Three removal types:** Remove a single file, the files in a folder (optionally recursive) or all empty folders below a path.
- ⏳ **Age based:** Only remove items older than a number of days, months or years, based on their creation date. Use `0` to remove everything.
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

Copy `Example.json` to `MyConfig.json` and adjust it. Every property is explained in the `?` section of `Example.json`. A minimal example:

```json
{
  "MaxConcurrent": {
    "JobsTotal": 4,
    "JobsPerComputer": 2
  },
  "Remove": {
    "FilesInFolder": [
      {
        "Name": "Application logs",
        "ComputerName": "SERVER1",
        "Path": "D:\\Logs",
        "Recurse": true,
        "OlderThan": {
          "Quantity": 30,
          "Unit": "Day"
        }
      }
    ],
    "EmptyFolders": [
      {
        "Name": "Drop zone",
        "ComputerName": null,
        "Path": "\\\\contoso\\share\\dropzone"
      }
    ]
  },
  "Settings": {
    "ScriptName": "Remove old logs",
    "SendMail": {
      "When": "OnError",
      "To": ["admin@example.com"]
    }
  }
}
```

`EmptyFolders` jobs always run after all other removal jobs, so folders emptied by those jobs are removed in the same run.

The mail server, log folder and event log settings are left out above for brevity; they are required, see `Example.json`.

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
