# Description
A PowerShell script designed to execute robocopy. 

This script will execute robocopy on a source and destination folder/file. All required parameters are read from a [.JSON file](Example.json). The input file can be modified as needed. 

A scheduled task can be created to run the script in an automated way.

# How does it work?
- Robocopy is executed and the results are stored in robocopy native log files.
- An e-mail is sent containing an overview and a link to the log files.

# Best practices
- Make sure to use the correct [robocopy arguments](https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy#parameters) as the switch '/MIR' can also delete files/folders.
- Read the [help section](Robocopy.ps1) in the script for an explanation of the input file parameters.

# Requirements
All required modules can be found in the [repository](https://github.com/DarkLite1?tab=repositories) and are free to use.