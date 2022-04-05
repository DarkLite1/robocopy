#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
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
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
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
            It 'MailTo is missing' {
                @{
                    # MailTo       = @('bob@contoso.com')
                    RobocopyTasks = @()
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'MailTo' addresses found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'RobocopyTasks is missing' {
                @{
                    MailTo = @('bob@contoso.com')
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
                        MailTo        = @('bob@contoso.com')
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
                        MailTo        = @('bob@contoso.com')
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
                        MailTo        = @('bob@contoso.com')
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
                It 'Source is a local path but no ComputerName is given' {
                    @{
                        MailTo        = @('bob@contoso.com')
                        RobocopyTasks = @(
                            @{
                                Name         = $null
                                Source       = 'd:\bla'
                                Destination  = '\\x:\b'
                                Switches     = '/x /y /c'
                                File         = $null
                                ComputerName = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*source 'd:\bla' and destination '\\x:\b': No 'ComputerName' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Destination is a local path but no ComputerName is given' {
                    @{
                        MailTo        = @('bob@contoso.com')
                        RobocopyTasks = @(
                            @{
                                Name         = $null
                                Source       = '\\x:\b'
                                Destination  = 'd:\bla'
                                Switches     = '/x /y /c'
                                File         = $null
                                ComputerName = $null
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                    
                    .$testScript @testParams
                    
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*source '\\x:\b' and destination 'd:\bla': No 'ComputerName' found*")
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
            MailTo        = @('bob@contoso.com')
            RobocopyTasks = @(
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
        It 'to the user in MailTo' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                $MailTo -eq 'bob@contoso.com'
            }
        }
        It 'with a summary of the copied data' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($MailTo -eq 'bob@contoso.com') -and
                ($Message -like "*<a href=`"\\$ENV:COMPUTERNAME\*source`">\\$ENV:COMPUTERNAME\*source</a><br>*<a href=`"\\$ENV:COMPUTERNAME\*destination`">\\$ENV:COMPUTERNAME\*destination</a>*")
            }
        }
    }
} -Tag test