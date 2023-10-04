#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
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
            Context 'SendMail' {
                It 'To is missing' {
                    @{
                        SendMail      = @{
                            Header = $null
                            # To     = @('bob@contoso.com')
                            When   = 'Always'
                        }
                        RobocopyTasks = @()
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'SendMail.To' addresses found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'When is missing' {
                    @{
                        SendMail      = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            # When   = 'Always'
                        }
                        RobocopyTasks = @()
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'SendMail.When' found, valid options are: Never, OnlyOnError, OnlyOnErrorOrCopies or Always.")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'When contains an invalid value' {
                    @{
                        SendMail      = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            When   = 'Wrong'
                        }
                        RobocopyTasks = @()
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'Wrong' in 'SendMail.When' is not valid, valid options are: Never, OnlyOnError, OnlyOnErrorOrCopies or Always.")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            It 'RobocopyTasks is missing' {
                @{
                    SendMail = @{
                        Header = $null
                        To     = @('bob@contoso.com')
                        When   = 'Always'
                    }
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'RobocopyTasks' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'RobocopyTasks' {
                It 'Source is missing' {
                    @{
                        SendMail      = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            When   = 'Always'
                        }
                        RobocopyTasks = @(
                            @{
                                Name         = $null
                                # Source       = '\\x:\a'
                                Destination  = '\\x:\b'
                                Switches     = '/x /y /c'
                                File         = $null
                                ComputerName = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Source' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Destination is missing' {
                    @{
                        SendMail      = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            When   = 'Always'
                        }
                        RobocopyTasks = @(
                            @{
                                Name         = $null
                                Source       = '\\x:\a'
                                # Destination  = '\\x:\b'
                                Switches     = '/x /y /c'
                                File         = $null
                                ComputerName = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Destination' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Switches is missing' {
                    @{
                        SendMail      = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            When   = 'Always'
                        }
                        RobocopyTasks = @(
                            @{
                                Name         = $null
                                Source       = '\\x:\a'
                                Destination  = '\\x:\b'
                                # Switches     = '/x /y /c'
                                File         = $null
                                ComputerName = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Switches' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'ComputerName is used together with UNC paths (double hop issue)' {
                    @{
                        SendMail      = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            When   = 'Always'
                        }
                        RobocopyTasks = @(
                            @{
                                Name         = $null
                                Source       = '\\x$\b'
                                Destination  = '\\x$\c'
                                Switches     = '/x /y /c'
                                ComputerName = $env:COMPUTERNAME
                                File         = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile' ComputerName '$env:COMPUTERNAME', Source '\\x$\b', Destination '\\x$\c': When ComputerName is used only local paths are allowed. This to avoid the double hop issue*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'ComputerName is not used together with a local path' {
                    @{
                        SendMail      = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            When   = 'Always'
                        }
                        RobocopyTasks = @(
                            @{
                                Name         = $null
                                Source       = 'x:\b'
                                Destination  = '\\x$\c'
                                Switches     = '/x /y /c'
                                ComputerName = $null
                                File         = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile' Source 'x:\b', Destination '\\x$\c': When ComputerName is not used only UNC paths are allowed.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'MaxConcurrentJobs' {
                It 'is missing' {
                    @{
                        SendMail      = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            When   = 'Always'
                        }
                        # MaxConcurrentJobs = 2
                        RobocopyTasks = @(
                            @{
                                Name         = $null
                                Source       = '\\x:\a'
                                Destination  = '\\x:\b'
                                Switches     = '/x /y /c'
                                File         = $null
                                ComputerName = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

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
                    @{
                        SendMail          = @{
                            Header = $null
                            To     = @('bob@contoso.com')
                            When   = 'Always'
                        }
                        MaxConcurrentJobs = 'a'
                        RobocopyTasks     = @(
                            @{
                                Name         = $null
                                Source       = '\\x:\a'
                                Destination  = '\\x:\b'
                                Switches     = '/x /y /c'
                                File         = $null
                                ComputerName = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
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
            @{Path = 'source/sub'; Type = 'Container' }
            @{Path = 'source/sub/test'; Type = 'File' }
            @{Path = 'destination'; Type = 'Container' }
        ) | ForEach-Object {
            (New-Item "TestDrive:/$($_.Path)" -ItemType $_.Type).FullName
        }

        @{
            SendMail          = @{
                Header = $null
                To     = @('bob@contoso.com')
                When   = 'Always'
            }
            MaxConcurrentJobs = 2
            RobocopyTasks     = @(
                @{
                    Name         = $null
                    Source       = $testData[0]
                    Destination  = $testData[3]
                    Switches     = '/MIR /Z /NP /MT:8 /ZB'
                    File         = $null
                    ComputerName = $env:COMPUTERNAME
                }
            )
        } | ConvertTo-Json | Out-File @testOutParams
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