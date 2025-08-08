#Requires -Modules Pester
#Requires -Version 7

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
                        File        = $null
                        Switches    = '/COPY'
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
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
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
        Should -BeTrue
    }
}
Describe 'create an error log file when' {
    It 'the log folder cannot be created' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = 'x:\notExistingLocation'

        Test-NewJsonFileHC

        Mock Out-File

        .$testScript @testParams

        $LASTEXITCODE | Should -Be 1

        Should -Not -Invoke Out-File
    }
    Context 'the ConfigurationJsonFile' {
        It 'is not found' {
            Mock Out-File

            $testNewParams = $testParams.clone()
            $testNewParams.ConfigurationJsonFile = 'nonExisting.json'

            .$testScript @testNewParams

            $LASTEXITCODE | Should -Be 1

            Should -Not -Invoke Out-File
        }
        Context 'property' {
            It 'Tasks.<_> not found' -ForEach @(
                'Robocopy'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].$_ = $null

                Test-NewJsonFileHC

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should -BeLike "* Property 'Tasks.Robocopy.Arguments' or 'Tasks.Robocopy.InputFile' not found*"
            }
            It 'Tasks.Robocopy.Arguments.<_> not found' -ForEach @(
                'Source', 'Destination', 'Switches'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Robocopy.Arguments.$_ = $null

                Test-NewJsonFileHC

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                $testLogFileContent = Test-GetLogFileDataHC

                $testLogFileContent[0].Message |
                Should -BeLike "*Property 'Tasks.Robocopy.Arguments.$_' not found*"
            }
        }
    }
}
Describe 'when all tests pass with' {
    Describe 'Robocopy.Arguments' {
        BeforeAll {
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
                Switches    = '/MIR /Z /NP /MT:8 /ZB'
                File        = $null
            }

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'robocopy is executed' {
            @(
                "TestDrive:/destination",
                "TestDrive:/destination/sub/test"
            ) | Should -Exist
        }
        Context 'create a robocopy log file' {
            It 'in the log folder with the TaskName' {
                Get-ChildItem -Path $testInputFile.Settings.SaveLogFiles.Where.Folder -Filter '* - Test (Brecht) (Test) - name of the task (1) - Log.txt' | 
                Should -Not -BeNullOrEmpty
            }
        }
        Context 'send an e-mail' {
            It 'with attachment to the user' {
                Should -Invoke Send-MailKitMessageHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($From -eq 'm@example.com') -and
                    ($To -eq '007@example.com') -and
                    ($SmtpPort -eq 25) -and
                    ($SmtpServerName -eq 'SMTP_SERVER') -and
                    ($SmtpConnectionType -eq 'StartTls') -and
                    ($Subject -eq '1 job, 1 item, Email subject') -and
                    ($Credential) -and
                    ($Attachments -like '*- Log.txt') -and
                    # ($Body -like "*<a href=`"\\$ENV:COMPUTERNAME\*source`">\\$ENV:COMPUTERNAME\*source</a><br>*<a href=`"\\$ENV:COMPUTERNAME\*destination`">\\$ENV:COMPUTERNAME\*destination</a>*") -and
                    ($Body -like "*name of the task*") -and
                    ($MailKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll') -and
                    ($MimeKitAssemblyPath -eq 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll')
                }
            }
        } -Tag test
    }
    Describe 'Robocopy.FileInput' {
        BeforeAll {
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
                "TestDrive:/destination",
                "TestDrive:/destination/sub/test"
            ) | Should -Exist
        }
        Context 'create a robocopy log file' {
            It 'in the log folder with the name of the robocopy input file' {
                Get-ChildItem -Path $testInputFile.Settings.SaveLogFiles.Where.Folder -Filter '* - Test (Brecht) (Test) - RobocopyConfig.RCJ (1) - Log.txt' | 
                Should -Not -BeNullOrEmpty
            }
        }
        Context 'a mail is sent' {
            It 'to the user in SendMail.To' {
                Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                    $To -eq '007@example.com'
                }
            }
            It 'with a summary of the copied data' {
                Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                    ($To -eq '007@example.com') -and
                    ($Message -like "*<a href=`"$testRobocopyConfigFilePath`">$testRobocopyConfigFilePath</a>*")
                }
            }
        }
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
                        Switches    = '/MIR /Z /NP /MT:8 /ZB'
                        File        = $null
                    }
                }
            }
        }

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'execute Robocopy.exe with /MIR switch' {
        It 'source data is still present' {
            $testSourceData | ForEach-Object {
                $_ | Should -Exist
            }
        }
        It 'destination data is created' {
            foreach ($testDestFolder in $testDestinationFolder) {
                foreach ($testSrcData in $testSourceData) {
                    $testDestFolder + ($testSrcData -split 'source')[1] |
                    Should -Exist
                }
            }
        }
    }
    Context 'a mail is sent' {
        It 'to the user in SendMail.To' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                $To -eq '007@example.com'
            }
        }
        It 'with no error in Message' {
            Should -Not -Invoke Send-MailHC -Scope Describe -ParameterFilter {
                ($Message -Like "*System error*")
            }
        }
    }
}