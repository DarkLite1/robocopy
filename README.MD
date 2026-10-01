# 🚀 PowerShell Robocopy Automation

> **The ultimate wrapper for enterprise-grade file synchronization and backups.**

Tired of maintaining dozens of scattered batch scripts? This PowerShell script automates and manages your Robocopy tasks using a single, centralized JSON configuration file. With built-in support for parallel execution, comprehensive logging, and email alerts, it’s the perfect engine for scheduled data synchronization and reliable backups.

---

## ✨ Why use this?

- ⚡ **Lightning Fast:** Run multiple Robocopy tasks concurrently to slash backup times and optimize large-scale file operations.
- ⚙️ **Centralized Configuration:** Define all tasks, arguments, and settings in one clean, readable JSON file.
- 🌍 **Remote Ready:** Execute commands on local or remote machines seamlessly using PowerShell remoting.
- 📊 **Smart Logging:** Never guess if a backup failed. Route detailed execution and error logs to files or directly to the Windows Event Log.
- 📬 **Email Alerts:** Get automated summaries of executed tasks, errors, and actions straight to your inbox.
- 🔐 **Secure & Portable:** Inject environment variables for sensitive data (like passwords or server names) to keep your configs safe and portable.
- 🛠️ **Flexible Tasks:** Pass standard Robocopy arguments or reference your pre-existing Robocopy job files (`.RCJ`).

---

## 🚀 Getting Started

### 1. Prerequisites

- **PowerShell 7.x+**: Built for modern PowerShell.
- **Robocopy.exe**: Native to Windows, no installation required.
- **Email Dependencies (Optional)**: If you want email alerts, install the MailKit and MimeKit modules via an elevated PowerShell prompt:
  ```powershell
  Install-Package -Name 'MailKit','MimeKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  ```

### 2. Installation

Clone the repository to your local machine:

```bash
git clone https://github.com/DarkLite1/robocopy.git
cd robocopy
```

### 3. Configuration

The magic happens in the JSON file. Duplicate `Example.json` and rename it to `MyConfig.json`. Here is a quick look at how easy it is to define a task:

```json
{
  "Tasks": [
    {
      "TaskName": "Nightly_Share_Backup",
      "ComputerName": null,
      "Robocopy": {
        "InputFile": null,
        "Arguments": {
          "Source": "C:\\LocalData\\Shares",
          "Destination": "\\\\NAS\\Backups\\Shares",
          "Files": [],
          "Switches": ["/MIR", "/R:3", "/W:5", "/NP"]
        }
      }
    }
  ],
  "Settings": {
    "SendMail": {
      "When": "OnError",
      "To": "admin@yourdomain.com"
    }
  }
}
```

## 💻 Usage

To run your backup jobs manually, execute the script and point it to your JSON configuration:

```powershell
& '.\Robocopy.ps1' -ConfigurationJsonFile '.\MyConfig.json'
```

## ⏱️ Automating with Task Scheduler

For hands-off, recurring backups, set this up in Windows Task Scheduler. Use pwsh.exe to invoke the script and pass the exit code back to Windows so your task history accurately reflects success or failure.

- Program: `pwsh.exe`
- Arguments:

```
-Command  "& 'C:\Robocopy\Robocopy.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit $LASTEXITCODE"
```

**_Pro Tip:_** Run the task with highest privileges and ensure the execution account has read/write access to all source and destination paths.

## 🚦 Exit Codes

This script standardizes Robocopy's notorious exit codes into simple binary outcomes, making it easy to monitor via third-party tools:

| Code | Status  | Description                                                                                                                                                                     |
| ---- | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0    | Success | All tasks completed successfully. (Includes native Robocopy codes 0–7: no change, copied, extra, or mismatch).                                                                  |
| 1    | Error   | Something went wrong. This triggers on an invalid config file, a task failing to start, a native Robocopy error (code 8+), or a script failure like an email/event log timeout. |
