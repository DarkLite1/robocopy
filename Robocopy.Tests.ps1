#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testInputFile = @{
        MaxConcurrentJobs = 1
        SendMail          = @{
            To   = 'bob@contoso.com'
            When = 'Always'
        }
        Tasks             = @(
            @{
                Name         = 'Copy files'
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
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'MaxConcurrentJobs', 'Tasks', 'SendMail'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'SendMail' {
                It 'SendMail.<_> not found' -ForEach @(
                    'To', 'When'
                ) {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.SendMail.$_ = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'SendMail.When is not valid' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.SendMail.When = 'wrong'

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'Tasks' {
                Context 'Tasks.ComputerName' {
                    It 'ComputerName is used with UNC paths (double hop issue)' {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME
                        $testNewInputFile.Tasks[0].Robocopy.Arguments.Source = '\\x$\b'
                        $testNewInputFile.Tasks[0].Robocopy.Arguments.Destination = '\\x$\c'

                        $testNewInputFile | ConvertTo-Json -Depth 7 |
                        Out-File @testOutParams

                        .$testScript @testParams

                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and ($Message -like "*ComputerName '$env:COMPUTERNAME', Source '\\x$\b', Destination '\\x$\c': When ComputerName is used only local paths are allowed. This to avoid the double hop issue*")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'ComputerName is not used with a local path' {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Tasks[0].ComputerName = $null
                        $testNewInputFile.Tasks[0].Robocopy.Arguments.Source = 'x:\b'
                        $testNewInputFile.Tasks[0].Robocopy.Arguments.Destination = '\\x$\c'

                        $testNewInputFile | ConvertTo-Json -Depth 7 |
                        Out-File @testOutParams

                        .$testScript @testParams

                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                            (&$MailAdminParams) -and ($Message -like "*Source 'x:\b', Destination '\\x$\c': When ComputerName is not used only UNC paths are allowed.*")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
                Context 'Tasks.Robocopy.Arguments' {
                    It 'Tasks.Robocopy.Arguments.<_> not found' -ForEach @(
                        'Source', 'Destination', 'Switches'
                    ) {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Tasks[0].Robocopy.Arguments.$_ = $null

                        $testNewInputFile | ConvertTo-Json -Depth 7 |
                        Out-File @testOutParams

                        .$testScript @testParams

                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Tasks.Robocopy.Arguments.$_' not found*")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
                It 'contains no Robocopy.Arguments or Robocopy.InputFile' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Robocopy.Arguments = $null
                    $testNewInputFile.Tasks[0].Robocopy.InputFile = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Tasks.Robocopy.Arguments' or 'Tasks.Robocopy.InputFile' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'contains both Robocopy.Arguments and Robocopy.InputFile' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Robocopy.InputFile = $testOutParams.FilePath

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Tasks.Robocopy.Arguments' and 'Tasks.Robocopy.InputFile' cannot be used at the same time*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'MaxConcurrentJobs' {
                It 'is missing' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.MaxConcurrentJobs = $null

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not a number' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.MaxConcurrentJobs = 'a'

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' needs to be a number, the value 'a' is not supported*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe 'when all tests pass' {
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
        $testNewInputFile.MaxConcurrentJobs = 2
        $testNewInputFile.Tasks[0].Name = $null
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
    Context 'a mail is sent' {
        It 'to the user in SendMail.To' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                $To -eq 'bob@contoso.com'
            }
        }
        It 'with a summary of the copied data' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($To -eq 'bob@contoso.com') -and
                ($Message -like "*<a href=`"\\$ENV:COMPUTERNAME\*source`">\\$ENV:COMPUTERNAME\*source</a><br>*<a href=`"\\$ENV:COMPUTERNAME\*destination`">\\$ENV:COMPUTERNAME\*destination</a>*")
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
        $testNewInputFile.MaxConcurrentJobs = 6
        $testNewInputFile.Tasks = $testDestinationFolder | ForEach-Object {
            @{
                Name         = $null
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
                $To -eq 'bob@contoso.com'
            }
        }
        It 'with no error in Message' {
            Should -Not -Invoke Send-MailHC -Scope Describe -ParameterFilter {
                ($Message -Like "*System error*")
            }
        }
    }
}