# PowerShell Robocopy Automation

A PowerShell script to automate and manage Robocopy tasks using a centralized JSON configuration file. It supports parallel execution, configurable logging, and email notifications, making it ideal for scheduled data synchronization and backup jobs.

## 🌟 Features

- **Concurrent Execution**: Run multiple Robocopy tasks simultaneously to optimize large-scale file operations.
- **Centralized Configuration**: All tasks and settings are defined in a single, easy-to-read JSON file.
- **Flexible Task Definition**: Configure tasks using standard Robocopy arguments or by referencing a pre-existing Robocopy job file (.RCJ).
- **Remote Operation**: Execute Robocopy commands on local or remote machines using PowerShell remoting.
- **Comprehensive Logging**: Save detailed Robocopy and script error logs to files or the Windows Event Log.
- **Email Notifications**: Send a summary of all executed tasks, errors, and actions via email, with options for when to send.
- **Environment Variable Support**: Use environment variables for sensitive information like passwords or server names, enhancing security and portability.

## 🚀 Getting Started

### Prerequisites

- **PowerShell 7.x or later**: This script is built for modern PowerShell.
- **Robocopy.exe**: The executable `robocopy.exe`, available on all Windows operating systems.
- **Emailing Prerequisites (Optional)**: If you plan to use email notifications, install the `MailKit` and `MimeKit` NuGet packages.
  ```powershell
  Install-Package -Name 'MailKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  Install-Package -Name 'MimeKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  ```

### Installation

1.  Clone this repository to your local machine:
    ```bash
    git clone https://github.com/DarkLite1/robocopy.git
    cd robocopy
    ```
2.  Review `Example.json` to understand the configuration options.
3.  Create your own configuration file (e.g., `MyConfig.json`) based on `Example.json`.

## 📝 Usage

Execute the script from PowerShell, providing the path to your configuration JSON file.

```powershell
& '.\Robocopy.ps1' -ConfigurationJsonFile '.\MyConfig.json'
```

### Scheduled Task Example ⏱️

For automated, recurring executions (e.g., via Windows Task Scheduler), you'll typically invoke the script using `pwsh.exe -Command`. This allows you to pass parameters directly.

- Program: `pwsh.exe`
- Arguments:

```
-Command  "& 'C:\Robocopy\Robocopy.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit $LASTEXITCODE"
```
