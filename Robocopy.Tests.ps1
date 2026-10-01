#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.2.0' }
#Requires -Version 7

BeforeDiscovery {
    $testRemotingAvailable = try {
        $null = Invoke-Command -ComputerName 127.0.0.1 -ConfigurationName PowerShell.7 -ScriptBlock { 1 } -ErrorAction Stop
        $true
    }
    catch {
        $false
    }
}

BeforeAll {
    $testInputFile = @{
        MaxConcurrentTasks = 1
        Tasks              = @(
            @{
                TaskName     = 'Copy files'
                ComputerName = 'PC1'
                Robocopy     = @{
                    InputFile = $null
                    Arguments = @{
                        Source      = 'TestDrive:\source'
                        Destination = 'TestDrive:\destination'
                        Files       = @()
                        Switches    = @('/COPY')
                    }
                }
            }
        )
        Settings           = @{
            ScriptName     = 'Test (Brecht)'
            SendMail       = @{
                When         = 'Always'
                From         = 'm@example.com'
                To           = '007@example.com'
                Subject      = 'Email subject'
                Body         = 'Email body'
                Smtp         = @{
                    ServerName     = 'SMTP_SERVER'
                    Port           = 25
                    ConnectionType = 'StartTls'
                    UserName       = 'bob'
                    Password       = 'pass'
                }
                AssemblyPath = @{
                    MailKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }
            }
            SaveLogFiles   = @{
                What                = @{
                    SystemErrors = $true
                    RobocopyLogs = $true
                }
                Where               = @{
                    Folder = (New-Item 'TestDrive:/log' -ItemType Directory).FullName
                }
                DeleteLogsAfterDays = 1
            }
            SaveInEventLog = @{
                Save    = $true
                LogName = 'Scripts'
            }
        }
    }

    $testOutParams = @{
        FilePath = (New-Item 'TestDrive:/Test.json' -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ConfigurationJsonFile = $testOutParams.FilePath
    }

    function Copy-ObjectHC {
        <#
        .SYNOPSIS
            Make a deep copy of an object using JSON serialization.

        .DESCRIPTION
            Uses ConvertTo-Json and ConvertFrom-Json to create an independent
            copy of an object. This method is generally effective for objects
            that can be represented in JSON format.

        .PARAMETER InputObject
            The object to copy.

        .EXAMPLE
            $newArray = Copy-ObjectHC -InputObject $originalArray
        #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $jsonString = $InputObject | ConvertTo-Json -Depth 100

        $deepCopy = $jsonString | ConvertFrom-Json

        return $deepCopy
    }

    function Test-GetLogFileDataHC {
        param (
            [String]$FileNameRegex = '* - System errors log.json',
            [String]$LogFolderPath = $testInputFile.Settings.SaveLogFiles.Where.Folder
        )

        $testLogFile = Get-ChildItem -Path $LogFolderPath -File -Filter $FileNameRegex

        if ($testLogFile.count -eq 1) {
            Get-Content $testLogFile | ConvertFrom-Json
        }
        elseif (-not $testLogFile) {
            throw "No log file found in folder '$LogFolderPath' matching '$FileNameRegex'"
        }
        else {
            throw "Found multiple log files in folder '$LogFolderPath' matching '$FileNameRegex'"
        }
    }

    function Send-MailKitMessageHC {
        param (
            [parameter(Mandatory)]
            [string]$MailKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$MimeKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$SmtpServerName,
            [parameter(Mandatory)]
            [ValidateSet(25, 465, 587, 2525)]
            [int]$SmtpPort,
            [parameter(Mandatory)]
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
            [string[]]$To,
            [string[]]$Bcc,
            [int]$MaxAttachmentSize = 20MB,
            [ValidateSet(
                'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
            )]
            [string]$SmtpConnectionType = 'None',
            [ValidateSet('Normal', 'Low', 'High')]
            [string]$Priority = 'Normal',
            [string[]]$Attachments,
            [PSCredential]$Credential
        )
    }

    function Test-NewJsonFileHC {
        try {
            if (-not $testNewInputFile) {
                throw "Variable '$testNewInputFile' cannot be blank"
            }

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams
        }
        catch {
            throw "Failure in Test-NewJsonFileHC: $_"
        }
    }

    Mock Send-MailKitMessageHC
    Mock New-EventLog
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ConfigurationJsonFile') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should-BeTrue
    }
}
Describe 'create an error log file when' {
    It 'the log folder cannot be created' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = 'x:\notExistingLocation'

        Test-NewJsonFileHC

        Mock Out-File

        .$testScript @testParams

        $LASTEXITCODE | Should-Be 1

        Should-NotInvoke Out-File
    }
    Context 'the ConfigurationJsonFile' {
        It 'is not found' {
            Mock Out-File

            $testNewParams = $testParams.clone()
            $testNewParams.ConfigurationJsonFile = 'nonExisting.json'

            .$testScript @testNewParams

            $LASTEXITCODE | Should-Be 1

            Should-NotInvoke Out-File
        }
        Context 'property' {
            It 'Tasks.<_> not found' -ForEach @(
                'Robocopy'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].$_ = $null

                Test-NewJsonFileHC

                .$testScript @testParams

                $LASTEXITCODE | Should-Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should-BeLikeString "* Property 'Tasks.Robocopy.Arguments' or 'Tasks.Robocopy.InputFile' not found*"
            }
            It 'Tasks.Robocopy.Arguments.<_> not found' -ForEach @(
                'Source', 'Destination', 'Switches'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Robocopy.Arguments.$_ = $null

                Test-NewJsonFileHC

                .$testScript @testParams

                $LASTEXITCODE | Should-Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should-BeLikeString "*Property 'Tasks.Robocopy.Arguments.$_' not found*"
            }
            It 'Tasks.Robocopy.Arguments.<_> not found in the second task' -ForEach @(
                'Source', 'Destination', 'Switches'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks = @(
                    $testNewInputFile.Tasks[0],
                    (Copy-ObjectHC $testInputFile.Tasks[0])
                )
                $testNewInputFile.Tasks[1].Robocopy.Arguments.$_ = $null

                Test-NewJsonFileHC

                Get-ChildItem -Path $testInputFile.Settings.SaveLogFiles.Where.Folder -Filter '* - System errors log.json' |
                Remove-Item

                .$testScript @testParams

                $LASTEXITCODE | Should-Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should-BeLikeString "*Property 'Tasks.Robocopy.Arguments.$_' not found*"
            }
        }
    }
}
Describe 'when all tests pass with' {
    Describe 'Robocopy.Arguments' {
        BeforeAll {
            $Error.Clear()

            $testData = @(
                @{Path = 'source'; Type = 'Container' }
                @{Path = 'source\sub'; Type = 'Container' }
                @{Path = 'source\sub\test'; Type = 'File' }
                @{Path = 'destination'; Type = 'Container' }
            ) | ForEach-Object {
                (New-Item "TestDrive:\$($_.Path)" -ItemType $_.Type).FullName
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.MaxConcurrentTasks = 2
            $testNewInputFile.Tasks[0].TaskName = 'name of the task'
            $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME
            $testNewInputFile.Tasks[0].Robocopy.Arguments = @{
                Source      = $testData[0]
                Destination = $testData[3]
                Switches    = @('/MIR', '/Z', '/NP', '/MT:8', '/ZB')
                Files       = @()
            }

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'robocopy is executed' {
            @(
                'TestDrive:/destination',
                'TestDrive:/destination/sub/test'
            ) | Should-All { Test-Path -Path $_ | Should-BeTrue }
        }
        Context 'create a robocopy log file' {
            It 'in the log folder with the TaskName' {
                Get-ChildItem -Path $testInputFile.Settings.SaveLogFiles.Where.Folder -Filter '* - Test (Brecht) (Test) - name of the task - Log.txt' |
                Should-NotBeNull
            }
        }
        Context 'send an e-mail' {
            It 'with attachment to the user' {
                Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                    ($From -eq 'm@example.com') -and
                    ($To -eq '007@example.com') -and
                    ($SmtpPort -eq 25) -and
                    ($SmtpServerName -eq 'SMTP_SERVER') -and
                    ($SmtpConnectionType -eq 'StartTls') -and
                    ($Subject -eq '1 task, 1 file, Email subject') -and
                    ($Credential) -and
                    ($Attachments -like '*- Log.txt') -and
                    # ($Body -like "*<a href=`"\\$ENV:COMPUTERNAME\*source`">\\$ENV:COMPUTERNAME\*source</a><br>*<a href=`"\\$ENV:COMPUTERNAME\*destination`">\\$ENV:COMPUTERNAME\*destination</a>*") -and
                    ($Body -like '*name of the task*') -and
                    ($MailKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll') -and
                    ($MimeKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll')
                }
            }
        }
    }
    Describe 'Robocopy.FileInput' {
        BeforeAll {
            $Error.Clear()

            $testData = @(
                @{Path = 'source'; Type = 'Container' }
                @{Path = 'source\sub'; Type = 'Container' }
                @{Path = 'source\sub\test'; Type = 'File' }
                @{Path = 'destination'; Type = 'Container' }
            ) | ForEach-Object {
                (New-Item "TestDrive:\$($_.Path)" -ItemType $_.Type).FullName
            }

            $testRobocopyConfigFilePath = 'TestDrive:\RobocopyConfig.RCJ'

            $testRobocopyConfigFile = @"
/SD:$($testData[0])\    :: Source Directory.
/DD:$($testData[3])\    :: Destination Directory.
/IF		:: Include Files matching these names
/XD		:: eXclude Directories matching these names
/XF		:: eXclude Files matching these names
/S		:: copy Subdirectories, but not empty ones.
/E		:: copy subdirectories, including Empty ones.
/DCOPY:DA	:: what to COPY for directories (default is /DCOPY:DA).
/COPY:DAT	:: what to COPY for files (default is /COPY:DAT).
/PURGE		:: delete dest files/dirs that no longer exist in source.
/MIR		:: MIRror a directory tree (equivalent to /E plus /PURGE).
/ZB		:: use restartable mode; if access denied use Backup mode.
/R:5		:: number of Retries on failed copies: default 1 million.
/W:30		:: Wait time between retries: default is 30 seconds.
/NP		:: No Progress - don't display percentage copied.
"@

            $testRobocopyConfigFile | Out-File -FilePath $testRobocopyConfigFilePath -Encoding utf8

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.MaxConcurrentTasks = 1
            $testNewInputFile.Tasks[0].TaskName = $null
            $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME
            $testNewInputFile.Tasks[0].Robocopy.Arguments = $null
            $testNewInputFile.Tasks[0].Robocopy.InputFile = $testRobocopyConfigFilePath

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'robocopy is executed' {
            @(
                'TestDrive:/destination',
                'TestDrive:/destination/sub/test'
            ) | Should-All { Test-Path -Path $_ | Should-BeTrue }
        }
        Context 'create a robocopy log file' {
            It 'in the log folder with the name of the robocopy input file' {
                Get-ChildItem -Path $testInputFile.Settings.SaveLogFiles.Where.Folder -Filter '* - Test (Brecht) (Test) - RobocopyConfig.RCJ - Log.txt' |
                Should-NotBeNull
            }
        }
        Context 'send an e-mail' {
            It 'with attachment to the user' {
                Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                    ($From -eq 'm@example.com') -and
                    ($To -eq '007@example.com') -and
                    ($SmtpPort -eq 25) -and
                    ($SmtpServerName -eq 'SMTP_SERVER') -and
                    ($SmtpConnectionType -eq 'StartTls') -and
                    ($Subject -eq '1 task, 1 file, Email subject') -and
                    ($Credential) -and
                    ($Attachments -like '*- Log.txt') -and
                    # ($Body -like "*<a href=`"\\$ENV:COMPUTERNAME\*source`">\\$ENV:COMPUTERNAME\*source</a><br>*<a href=`"\\$ENV:COMPUTERNAME\*destination`">\\$ENV:COMPUTERNAME\*destination</a>*") -and
                    ($Body -like "*<a href=`"$testRobocopyConfigFilePath`">$testRobocopyConfigFilePath</a>*") -and
                    ($MailKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll') -and
                    ($MimeKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll')
                }
            }
        }
    }
}
Describe 'Convert-RobocopyLogToObjectHC' {
    BeforeAll {
        $testAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $testScript, [ref]$null, [ref]$null
        )
        $testFunctionAst = $testAst.Find(
            {
                param($ast)
                ($ast -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and
                ($ast.Name -eq 'Convert-RobocopyLogToObjectHC')
            }, $true
        )
        . ([scriptblock]::Create($testFunctionAst.Extent.Text))
    }
    It 'a destination path that contains the word Source is not seen as the source' {
        $actual = Convert-RobocopyLogToObjectHC -LogContent @(
            '  Source : C:\Data\',
            '    Dest : \\server\Source\Backup\'
        )

        $actual.Source | Should-Be 'C:\Data\'
        $actual.Destination | Should-Be '\\server\Source\Backup\'
    }
}
Describe 'an input file used on a remote computer' -Skip:(-not $testRemotingAvailable) {
    BeforeAll {
        $testSource = (New-Item 'TestDrive:\remoteInputFile\source' -ItemType Directory).FullName
        $testDestination = (New-Item 'TestDrive:\remoteInputFile\destination' -ItemType Directory).FullName
        $null = New-Item "$testSource\file.txt" -ItemType File

        # a TestDrive path only exists in this session, not on the remote one
        $testRobocopyConfigFilePath = 'TestDrive:\remoteInputFile\Job.RCJ'
        @"
/SD:$testSource\
/DD:$testDestination\
/E
"@ | Out-File -FilePath $testRobocopyConfigFilePath -Encoding utf8

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = '127.0.0.1'
        $testNewInputFile.Tasks[0].Robocopy.Arguments = $null
        $testNewInputFile.Tasks[0].Robocopy.InputFile = $testRobocopyConfigFilePath

        Test-NewJsonFileHC

        .$testScript @testParams
    }
    It 'robocopy is executed with the content of the local input file' {
        "$testDestination\file.txt" | Should-All { Test-Path -LiteralPath $_ | Should-BeTrue }
    }
}
Describe 'parallel tasks with input files that have the same name' {
    BeforeAll {
        $testTasks = foreach ($name in 'A', 'B') {
            $source = (New-Item "TestDrive:\sameName$name\source" -ItemType Directory).FullName
            $destination = (New-Item "TestDrive:\sameName$name\destination" -ItemType Directory).FullName
            $null = New-Item "$source\file$name.txt" -ItemType File

            $inputFile = "$((Get-Item "TestDrive:\sameName$name").FullName)\Job.RCJ"
            @"
/SD:$source\
/DD:$destination\
/E
"@ | Out-File -FilePath $inputFile -Encoding utf8

            @{
                TaskName     = "Task $name"
                ComputerName = $env:COMPUTERNAME
                Robocopy     = @{
                    InputFile = $inputFile
                    Arguments = $null
                }
            }
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.MaxConcurrentTasks = 2
        $testNewInputFile.Tasks = @($testTasks)

        Test-NewJsonFileHC

        .$testScript @testParams
    }
    It 'each task copies its own files' {
        'TestDrive:\sameNameA\destination\fileA.txt' | Should-All { Test-Path -LiteralPath $_ | Should-BeTrue }
        'TestDrive:\sameNameB\destination\fileB.txt' | Should-All { Test-Path -LiteralPath $_ | Should-BeTrue }
        'TestDrive:\sameNameA\destination\fileB.txt' | Should-All { Test-Path -LiteralPath $_ | Should-BeFalse }
        'TestDrive:\sameNameB\destination\fileA.txt' | Should-All { Test-Path -LiteralPath $_ | Should-BeFalse }
    }
}
Describe 'when a task fails to start' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = 'PC1'
        $testNewInputFile.Settings.SaveLogFiles.What.SystemErrors = $false

        Test-NewJsonFileHC

        $global:LASTEXITCODE = 0

        .$testScript @testParams
    }
    It 'the script exits with error code 1' {
        $LASTEXITCODE | Should-Be 1
    }
    It 'the error is written to the event log' {
        Should-Invoke Write-EventLog -Scope Describe -ParameterFilter {
            ($EntryType -eq 'Error') -and
            ($Message -like "*Task 'Copy files' on 'PC1'*")
        }
    }
    It 'the error is reported in the e-mail' {
        Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Subject -eq '1 task, 0 files, 1 error, Email subject') -and
            ($Body -like '*<th>Job errors</th>*<td>1</td>*') -and
            ($Body -like '*Copy files*')
        }
    }
}
Describe 'when a task fails to start and SaveInEventLog.Save is false' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = 'PC1'
        $testNewInputFile.Settings.SaveInEventLog.Save = $false

        Test-NewJsonFileHC

        .$testScript @testParams
    }
    It 'nothing is written to the event log' {
        Should-NotInvoke Write-EventLog -Scope Describe
    }
}
Describe 'when writing to the event log fails' {
    BeforeAll {
        Mock Write-EventLog { throw 'Event log failure' }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Source = (New-Item 'TestDrive:\eventLogFailureSource' -ItemType Directory).FullName
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Destination = (New-Item 'TestDrive:\eventLogFailureDestination' -ItemType Directory).FullName
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Switches = @('/E')

        Test-NewJsonFileHC

        .$testScript @testParams
    }
    It 'the system error is counted in the e-mail' {
        Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Subject -eq '1 task, 0 files, 1 error, Email subject') -and
            ($Priority -eq 'High') -and
            ($Body -like '*<th>System errors</th>*<td>1</td>*')
        }
    }
}
Describe 'an incorrect Settings property in the input file' {
    BeforeAll {
        Mock Invoke-Command
    }
    It '<Description>' -ForEach @(
        @{
            Description = 'Settings.ScriptName missing'
            Change      = { param($s) $s.ScriptName = $null }
            Message     = "Property 'Settings.ScriptName' not found"
        }
        @{
            Description = 'Settings.SendMail.When missing'
            Change      = { param($s) $s.SendMail.When = $null }
            Message     = "Property 'Settings.SendMail.When' not found"
        }
        @{
            Description = 'Settings.SendMail.When not supported'
            Change      = { param($s) $s.SendMail.When = 'Sometimes' }
            Message     = "Property 'Settings.SendMail.When' with value 'Sometimes' is not supported*"
        }
        @{
            Description = 'Settings.SendMail.From missing'
            Change      = { param($s) $s.SendMail.From = $null }
            Message     = "Property 'Settings.SendMail.From' not found"
        }
        @{
            Description = 'Settings.SendMail.Smtp.ServerName missing'
            Change      = { param($s) $s.SendMail.Smtp.ServerName = $null }
            Message     = "Property 'Settings.SendMail.Smtp.ServerName' not found"
        }
        @{
            Description = 'Settings.SendMail.To and Bcc missing'
            Change      = { param($s) $s.SendMail.To = $null }
            Message     = "Property 'Settings.SendMail.To' or 'Settings.SendMail.Bcc' not found"
        }
        @{
            Description = 'Settings.SendMail.Smtp.Port not supported'
            Change      = { param($s) $s.SendMail.Smtp.Port = 26 }
            Message     = "Property 'Settings.SendMail.Smtp.Port' with value '26' is not supported*"
        }
        @{
            Description = 'Settings.SendMail.Smtp.ConnectionType not supported'
            Change      = { param($s) $s.SendMail.Smtp.ConnectionType = 'Wrong' }
            Message     = "Property 'Settings.SendMail.Smtp.ConnectionType' with value 'Wrong' is not supported*"
        }
        @{
            Description = 'Settings.SaveLogFiles.What.SystemErrors missing'
            Change      = { param($s) $s.SaveLogFiles.What.SystemErrors = $null }
            Message     = "Property 'Settings.SaveLogFiles.What.SystemErrors' not found"
        }
        @{
            Description = 'Settings.SaveLogFiles.What.RobocopyLogs not a boolean'
            Change      = { param($s) $s.SaveLogFiles.What.RobocopyLogs = 'yes' }
            Message     = "Property 'Settings.SaveLogFiles.What.RobocopyLogs' needs to be true or false, the value 'yes' is not supported."
        }
        @{
            Description = 'Settings.SaveLogFiles.DeleteLogsAfterDays not a number'
            Change      = { param($s) $s.SaveLogFiles.DeleteLogsAfterDays = 'abc' }
            Message     = "Property 'Settings.SaveLogFiles.DeleteLogsAfterDays' needs to be a positive number, the value 'abc' is not supported."
        }
        @{
            Description = 'Settings.SaveInEventLog missing'
            Change      = { param($s) $s.PSObject.Properties.Remove('SaveInEventLog') }
            Message     = "Property 'Settings.SaveInEventLog.Save' not found"
        }
        @{
            Description = 'Settings.SaveInEventLog.Save not a boolean'
            Change      = { param($s) $s.SaveInEventLog.Save = 'yes' }
            Message     = "Property 'Settings.SaveInEventLog.Save' needs to be true or false, the value 'yes' is not supported."
        }
        @{
            Description = 'Settings.SaveInEventLog.LogName missing when Save is true'
            Change      = { param($s) $s.SaveInEventLog.LogName = $null }
            Message     = "Property 'Settings.SaveInEventLog.LogName' not found"
        }
    ) {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        & $Change $testNewInputFile.Settings

        Test-NewJsonFileHC

        .$testScript @testParams -WarningVariable testWarnings -WarningAction SilentlyContinue

        $LASTEXITCODE | Should-Be 1

        ($testWarnings -join "`n") | Should-BeLikeString "*Input file '*': $Message*"

        Should-NotInvoke Invoke-Command -Scope It
    }
    It 'Settings.SendMail properties are not needed when SendMail.When is Never' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SendMail = [PSCustomObject]@{ When = 'Never' }

        Test-NewJsonFileHC

        .$testScript @testParams -WarningVariable testWarnings -WarningAction SilentlyContinue

        ($testWarnings -join "`n") | Should-NotBeLikeString '*Settings.SendMail*'

        Should-Invoke Invoke-Command -Times 1 -Exactly -Scope It
    }
}
Describe 'when the log folder cannot be created' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = 'x:\notExistingLocation'
        $testNewInputFile.Tasks[0].TaskName = 'task without log folder'
        $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Source = (New-Item 'TestDrive:\noLogSource' -ItemType Directory).FullName
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Destination = (New-Item 'TestDrive:\noLogDestination' -ItemType Directory).FullName
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Switches = @('/E')

        Test-NewJsonFileHC

        .$testScript @testParams
    }
    It 'the task is still reported in the e-mail' {
        Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
            $Body -like '*task without log folder*'
        }
    }
}
Describe 'a robocopy job that fails' {
    BeforeAll {
        $testRobocopyConfigFilePath = (New-Item 'TestDrive:\Failing.RCJ' -ItemType File).FullName

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME
        $testNewInputFile.Tasks[0].Robocopy.Arguments = $null
        $testNewInputFile.Tasks[0].Robocopy.InputFile = $testRobocopyConfigFilePath

        Test-NewJsonFileHC

        # the job copies the input file to $env:TEMP, a missing folder makes it fail
        $testOriginalTemp = $env:TEMP
        $env:TEMP = Join-Path $TestDrive 'notExisting'

        $global:LASTEXITCODE = 0

        try {
            .$testScript @testParams
        }
        finally {
            $env:TEMP = $testOriginalTemp
        }
    }
    It 'the script exits with error code 1' {
        $LASTEXITCODE | Should-Be 1
    }
    It 'the error is written to the event log' {
        Should-Invoke Write-EventLog -Scope Describe -ParameterFilter {
            ($EntryType -eq 'Error') -and
            ($Message -like "*Task 'Copy files' on '$env:COMPUTERNAME': Failed to create temp job file*")
        }
    }
    It 'is counted as one error' {
        Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
            $Subject -eq '1 task, 0 files, 1 error, Email subject'
        }
    }
    It 'is not reported as NO CHANGE in the e-mail' {
        Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Body -cnotlike '*NO CHANGE*') -and ($Body -clike '*>ERROR<br>*')
        }
    }
}
Describe 'a robocopy exit code of 8 or higher' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Source = (New-Item 'TestDrive:\fatalSource' -ItemType Directory).FullName
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Destination = (New-Item 'TestDrive:\fatalDestination' -ItemType Directory).FullName
        # '/COPY' without a value is an invalid parameter, robocopy exit code 16
        $testNewInputFile.Tasks[0].Robocopy.Arguments.Switches = @('/COPY')

        Test-NewJsonFileHC

        $global:LASTEXITCODE = 0

        .$testScript @testParams
    }
    It 'the script exits with error code 1' {
        $LASTEXITCODE | Should-Be 1
    }
    It 'the error is written to the event log' {
        Should-Invoke Write-EventLog -Scope Describe -ParameterFilter {
            ($EntryType -eq 'Error') -and
            ($Message -like "*Task 'Copy files' on '$env:COMPUTERNAME': Robocopy exit code 16 (FATAL ERROR)*")
        }
    }
}
Describe 'an input file with an empty Arguments object' {
    BeforeAll {
        $testSource = (New-Item 'TestDrive:\emptyArguments\source' -ItemType Directory).FullName
        $testDestination = (New-Item 'TestDrive:\emptyArguments\destination' -ItemType Directory).FullName
        $null = New-Item "$testSource\file.txt" -ItemType File

        $testRobocopyConfigFilePath = 'TestDrive:\emptyArguments\Job.RCJ'
        @"
/SD:$testSource\
/DD:$testDestination\
/E
"@ | Out-File -FilePath $testRobocopyConfigFilePath -Encoding utf8

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME
        $testNewInputFile.Tasks[0].Robocopy.Arguments = [PSCustomObject]@{}
        $testNewInputFile.Tasks[0].Robocopy.InputFile = $testRobocopyConfigFilePath

        Test-NewJsonFileHC

        $global:LASTEXITCODE = 0

        .$testScript @testParams
    }
    It 'robocopy is executed' {
        "$testDestination\file.txt" | Should-All { Test-Path -LiteralPath $_ | Should-BeTrue }
    }
    It 'the script exits without error code' {
        $LASTEXITCODE | Should-Be 0
    }
}
Describe 'tasks with the same TaskName' {
    BeforeAll {
        $testLogFolder = (New-Item 'TestDrive:\sameTaskNameLog' -ItemType Directory).FullName

        $testTasks = foreach ($name in 'A', 'B') {
            @{
                TaskName     = 'same name'
                ComputerName = $env:COMPUTERNAME
                Robocopy     = @{
                    InputFile = $null
                    Arguments = @{
                        Source      = (New-Item "TestDrive:\sameTaskName$name\source" -ItemType Directory).FullName
                        Destination = (New-Item "TestDrive:\sameTaskName$name\destination" -ItemType Directory).FullName
                        Switches    = @('/E')
                        Files       = @()
                    }
                }
            }
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @($testTasks)
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = $testLogFolder

        Test-NewJsonFileHC

        .$testScript @testParams
    }
    It 'each task gets its own robocopy log file' {
        @(Get-ChildItem -LiteralPath $testLogFolder -Filter '* - same name*- Log.txt').Count |
        Should-Be 2
    }
}
Describe 'stress test' {
    BeforeAll {
        $testSourceData = @(
            @{Path = 'folder'; Type = 'Container' }
            @{Path = 'folder\sub'; Type = 'Container' }
            @{Path = 'folder\sub\file'; Type = 'File' }
        ) | ForEach-Object {
            (New-Item "TestDrive:\source\$($_.Path)" -ItemType $_.Type).FullName
        }

        $testDestinationFolder = 1..20 | ForEach-Object {
            (New-Item "TestDrive:\destination\f$_" -ItemType 'Container').FullName
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.MaxConcurrentTasks = 6
        $testNewInputFile.Tasks = $testDestinationFolder | ForEach-Object {
            @{
                TaskName     = $null
                ComputerName = $env:COMPUTERNAME
                Robocopy     = @{
                    InputFile = $null
                    Arguments = @{
                        Source      = (Get-Item -Path 'TestDrive:\source').FullName
                        Destination = $_
                        Switches    = @('/MIR', '/Z', '/NP', '/MT:8', '/ZB')
                        Files       = @()
                    }
                }
            }
        }

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        $Error.Clear()

        .$testScript @testParams
    }
    Context 'execute Robocopy.exe with /MIR switch' {
        It 'source data is still present' {
            $testSourceData | Should-All { Test-Path -LiteralPath $_ | Should-BeTrue }
        }
        It 'destination data is created' {
            foreach ($testDestFolder in $testDestinationFolder) {
                foreach ($testSrcData in $testSourceData) {
                    Test-Path -LiteralPath ($testDestFolder + ($testSrcData -split 'source')[1]) |
                    Should-BeTrue
                }
            }
        }
    }
    Context 'send an e-mail' {
        It 'with no system errors and attachment to the user' {
            Should-Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($From -eq 'm@example.com') -and
                ($To -eq '007@example.com') -and
                ($SmtpPort -eq 25) -and
                ($SmtpServerName -eq 'SMTP_SERVER') -and
                ($SmtpConnectionType -eq 'StartTls') -and
                ($Subject -eq '20 tasks, 20 files, Email subject') -and
                ($Credential) -and
                ($Attachments -like '*- Log.txt') -and
                ($Body -notlike '*system errors*') -and
                ($MailKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll') -and
                ($MimeKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll')
            }
        }
    }
}