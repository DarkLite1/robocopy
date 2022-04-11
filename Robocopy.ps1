<#
    .SYNOPSIS
        Copy files and folders with Robocopy.exe

    .DESCRIPTION
        Copy files and folders with Robocopy.exe based on its advanced 
        parameters. The parameters will be read from the import file together 
        with the source and the destination folders.

    .PARAMETER MailTo
        E-mail addresses of where to send the summary e-mail

    .PARAMETER RobocopyTasks
        Collection of individual robocopy jobs

    .PARAMETER Name
        Display a name in the email message body instead of source and 
        destination paths

    .PARAMETER ComputerName
        The computer where to execute the robocopy executable. This allows
        for running robocopy on remote machines.

    .PARAMETER Source
        Source path, local or UNC, where to copy/move files from. First argument
        to robocopy.

    .PARAMETER Destination
        Destination path, local or UNC, where to copy/move files/folder too. Second argument to robocopy.

    .PARAMETER Switches
        Robocopy copy/move arguments, last argument to robocopy

    .PARAMETER File
        Robocopy file arguments, third argument to robocopy
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Int]$MaxConcurrentJobs = 4,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Robocopy\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    $scriptBlock = {
        Param (
            [Parameter(Mandatory)]
            [String]$Source,
            [Parameter(Mandatory)]
            [String]$Destination,
            [Parameter(Mandatory)]
            [String]$Switches,
            [String]$File,
            [String]$Name,
            [String]$ComputerName
        )

        Try {
            $result = [PSCustomObject]@{
                Name           = $Name
                ComputerName   = $ComputerName
                Source         = $Source
                Destination    = $Destination
                File           = $File
                Switches       = $Switches
                RobocopyOutput = $null
                ExitCode       = $null
                Error          = $null
            }
            
            $expression = [String]::Format(
                'ROBOCOPY "{0}" "{1}" {2} {3}', 
                $Source, $Destination, $File, $Switches
            )
            $result.robocopyOutput = Invoke-Expression $expression
        }
        Catch {
            $result.Error = $_
        }
        Finally {
            $result.ExitCode = $LASTEXITCODE
            $result
        }
    }

    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Date         = 'ScriptStartTime'
                NoFormatting = $true
                Unique       = $True
            }
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }
        if (-not ($RobocopyTasks = $file.RobocopyTasks)) {
            throw "Input file '$ImportFile': No 'RobocopyTasks' found."
        }
        foreach ($task in $RobocopyTasks) {
            if (-not $task.Source) {
                throw "Input file '$ImportFile': No 'Source' found in one of the 'RobocopyTasks'."
            }
            if (-not $task.Destination) {
                throw "Input file '$ImportFile': No 'Destination' found for source '$($task.Source)'."
            }
            if (
                (-not $task.ComputerName) -and
                (
                    ($task.Source -notMatch '^\\\\') -or 
                    ($task.Destination -notMatch '^\\\\')
                )
            ) {
                throw "Input file '$ImportFile' source '$($task.Source)' and destination '$($task.Destination)': No 'ComputerName' found."
            }
            if (-not $task.Switches) {
                throw "Input file '$ImportFile': No 'Switches' found for source '$($task.Source)'."
            }
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    #region Execute robocopy tasks
    $jobs = @()

    ForEach ($task in $RobocopyTasks) {
        $invokeParams = @{
            ScriptBlock  = $scriptBlock
            ArgumentList = $task.Source, $task.Destination, $task.Switches, $task.File, $task.Name, $task.ComputerName
        }

        $M = "Start job on '{0}' with Source '{1}' Destination '{2}' Switches '{3}' File '{4}' Name '{5}'" -f $(
            if ($task.ComputerName) { $task.ComputerName }
            else { $env:COMPUTERNAME }
        ),
        $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
        $invokeParams.ArgumentList[2], $invokeParams.ArgumentList[3], 
        $invokeParams.ArgumentList[4]
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        # & $scriptBlock -Source $invokeParams.ArgumentList[0] -destination $invokeParams.ArgumentList[1] -switches $invokeParams.ArgumentList[2]

        $jobs += if ($task.ComputerName) {
            $invokeParams.ComputerName = $task.ComputerName
            $invokeParams.AsJob = $true
            Invoke-Command @invokeParams
        }
        else {
            Start-Job @invokeParams
        }

        Wait-MaxRunningJobsHC -Name $jobs -MaxThreads $MaxConcurrentJobs
    }

    $M = "Wait for all $($jobs.count) jobs to be finished"
    Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

    $jobResults = if ($jobs) { 
        $jobs | Wait-Job -Force | Receive-Job 
    }
    #endregion
}

End {
    Try {
        $color = @{
            NoCopy   = 'White'     # Nothing copied
            CopyOk   = 'LightGrey' # Copy successful
            Mismatch = 'Orange'    # Clean-up needed (Mismatch)
            Fatal    = 'Red'       # Fatal error
        }

        $counter = @{
            totalItemsCopied    = 0
            robocopyBadExitCode = 0
            robocopyJobError    = 0
            systemError         = 0
        }
        
        $htmlTableRows = @() 

        Foreach ($job in $jobResults) {
            #region Get row color
            $rowColor = Switch ($job.ExitCode) {
                0 { 
                    $color.NoCopy 
                }
                { ($_ -ge 1) -and ($_ -le 3) } { 
                    $color.CopyOk 
                }
                { ($_ -ge 4) -and ($_ -le 7) } {
                    $color.Mismatch
                    $counter.robocopyBadExitCode++
                }
                default { 
                    $color.Fatal
                    $counter.robocopyBadExitCode++ 
                }
            }

            if ($job.Error) {
                $rowColor = $color.Fatal
                $counter.robocopyJobError++ 
            }
            #endregion

            #region Create robocopy log file
            $logParams.Name = $job.Destination + '.log'
            $logFile = New-LogFileNameHC @logParams
            $job.RobocopyOutput | Out-File -LiteralPath $logFile -Encoding utf8
            #endregion

            #region Convert robocopy log file
            $robocopyLogAnalyses = ConvertFrom-RobocopyLogHC -LogFile $logFile

            $itemsCopiedCount = [INT]$robocopyLogAnalyses.Files.Copied + 
            [INT]$robocopyLogAnalyses.Dirs.Copied
            $counter.totalItemsCopied += $itemsCopiedCount

            $robocopy = @{
                ExitMessage   = ConvertFrom-RobocopyExitCodeHC -ExitCode $job.ExitCode
                ExecutionTime = if (
                    $robocopyLogAnalyses.Times.Total
                ) {
                    $robocopyLogAnalyses.Times.Total 
                }
                else { 'NA' }
                ItemsCopied   = if ($itemsCopiedCount) { 
                    $itemsCopiedCount 
                }
                else { 'NA' }
            }
            #endregion

            #region Create HTML table rows
            $htmlTableRows += @"
<tr bgcolor="$rowColor" style="background:$rowColor;">
    <td id="TxtLeft">{0}<br>{1}{2}{3}</td>
    <td id="TxtLeft">$($robocopy.ExitMessage + ' (' + $job.ExitCode + ')')</td>
    <td id="TxtCentered">$($robocopy.ExecutionTime)</td>
    <td id="TxtCentered">$($robocopy.ItemsCopied)</td>
    <td id="TxtCentered">{4}</td>
</tr>
"@ -f 
            $(
                if ($job.Name) { 
                    $job.Name 
                }
                elseif ($job.Source -match '^\\\\') {
                    '<a href="{0}">{0}</a>' -f $job.Source
                }
                else {
                    $uncPath = $job.Source -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $job.ComputerName, $job.Source[0]
                    )
                    '<a href="{0}">{0}</a>' -f $uncPath
                }
            ),
            $(
                if ($job.Name) { 
                    $sourcePath = if ($job.Source -match '^\\\\') {
                        $job.Source
                    }
                    else {
                        $job.Source -Replace '^.{2}', (
                            '\\{0}\{1}$' -f $job.ComputerName, $job.Source[0]
                        )
                        '<a href="{0}">{0}</a>' -f $uncPath
                    }
                    $destinationPath = if ($job.Destination -match '^\\\\') {
                        $job.Destination
                    }
                    else {
                        $job.Destination -Replace '^.{2}', (
                            '\\{0}\{1}$' -f $job.ComputerName, $job.Destination[0]
                        )
                        '<a href="{0}">{0}</a>' -f $uncPath
                    }
                    '<a href="{0}">Source</a> > <a href="{1}">destination</a>' -f 
                    $sourcePath , $destinationPath 
                }
                elseif ($job.Destination -match '^\\\\') {
                    '<a href="{0}">{0}</a>' -f $job.Destination
                }
                else {
                    $uncPath = $job.Destination -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $job.ComputerName, $job.Destination[0]
                    )
                    '<a href="{0}">{0}</a>' -f $uncPath
                }
            ),
            $(
                if ($job.File) {
                    "<br>$($job.File)"
                }
            ),
            $(
                if ($job.Error) {
                    "<br><b>$($job.Error)</b>"
                }
            ),
            $(
                ConvertTo-HTMLlinkHC -Path $logFile -Name 'Log'
            )
            #endregion
        }

        #region Create HTML css
        $htmlCss = "
        <style>
            #TxtLeft{
                border: 1px solid Gray;
                border-collapse:collapse;
                text-align:left;
            }
            #TxtCentered {
                text-align: center;
                border: 1px solid Gray;
            }
            #LegendTable {
                border-collapse: collapse;
                table-layout: fixed;
                width: 600px;
            }
            #LegendRow {
                text-align: center;
                width: 150px;
                border: 1px solid Gray;
            }
        </style>"
        #endregion

        #region Create HTML table header
        $htmlTableHeaderRow = @"
            <tr>
                <th id="TxtLeft">Robocopy</th>
                <th id="TxtLeft">Message</th>
                <th id="TxtCentered" class="Centered">Total<br>time</th>
                <th id="TxtCentered" class="Centered">Copied<br>items</th>
                <th id="TxtCentered" class="Centered">Details</th>
            </tr>
"@
        #endregion

        #region Create HTML legend rows
        $htmlLegendRows = @"
    <tr>
        <td bgcolor="$($color.NoCopy)" style="background:$($color.NoCopy);" id="LegendRow">Nothing copied</td>
        <td bgcolor="$($color.CopyOk)" style="background:$($color.CopyOk);" id="LegendRow">Copy successful</td>
        <td bgcolor="$($color.Mismatch)" style="background:$($color.Mismatch);" id="LegendRow">Clean-up needed</td>
        <td bgcolor="$($color.Fatal)" style="background:$($color.Fatal);" id="LegendRow">Fatal error</td>
    </tr>
"@
        #endregion

        $logParams.Unique = $false
        $logParams.Name = ' - Mail.html'

        $mailParams = @{
            To        = $MailTo
            Priority  = 'Normal' 
            Subject   = '{0} jobs, {1} copied' -f $RobocopyTasks.Count, $counter.totalItemsCopied
            Message   = $null
            LogFolder = $LogFolder
            Header    = $ScriptName
            Save      = New-LogFileNameHC @logParams
        }
        
        #region Set mail subject and priority
        $uniqueSystemErrors = $Error.Exception.Message | 
        Where-Object { $_ } | Get-Unique

        $counter.systemError += $uniqueSystemErrors.Count

        if (
            $allErrorCount = $counter.systemError + 
            $counter.robocopyBadExitCode + $counter.robocopyJobError
        ) {
            $mailParams.Subject = "{0} error{1}, {2}" -f 
            $allErrorCount, $(if ($allErrorCount -ge 2) { 's' }), 
            $mailParams.Subject 
            $mailParams.Priority = 'High'
        }
        #endregion

        #region Create system errors HTML list
        $htmlUniqueSystemErrorsList = $null

        if ($uniqueSystemErrors) {
            $uniqueSystemErrors | ForEach-Object {
                Write-EventLog @EventErrorParams -Message $_
            }

            $htmlUniqueSystemErrorsList = $uniqueSystemErrors | 
            ConvertTo-HtmlListHC -Spacing Wide -Header 'System errors:'
        }
        #endregion

        #region Create HTML error overview table
        $htmlErrorOverviewTable = $null
        $htmlErrorOverviewTableRows = $null

        if ($counter.robocopyBadExitCode) {
            $htmlErrorOverviewTableRows += '<tr><th>{0}</th><td>{1}</td></tr>' -f $counter.robocopyBadExitCode, 'Errors in the robocopy log files'
        }
        if ($counter.robocopyJobError) {
            $htmlErrorOverviewTableRows += '<tr><th>{0}</th><td>{1}</td></tr>' -f $counter.robocopyJobError, 'Errors while executing robocopy'
        }
        if ($counter.systemError) {
            $htmlErrorOverviewTableRows += '<tr><th>{0}</th><td>{1}</td></tr>' -f $counter.systemError, 'System errors'
        }

        if ($htmlErrorOverviewTableRows) {
            $htmlErrorOverviewTable = "
            <p>Error overview:</p>
            <table>
                $htmlErrorOverviewTableRows
            </table>
            "
        }
        #endregion
        
        #region Create robocopy executed jobs table
        $htmlRobocopyExecutedJobsTable = $null

        if ($htmlTableRows) {
            $htmlRobocopyExecutedJobsTable = "
            <table id=`"TxtLeft`">
                $htmlTableHeaderRow
                $htmlTableRows
            </table>
            <br>
            <table id=`"LegendTable`">
                $htmlLegendRows
            </table>"    
        }
        #endregion
            
        $mailParams.Message = "
        $htmlCss
        $htmlErrorOverviewTable
        $htmlUniqueSystemErrorsList
        $htmlRobocopyExecutedJobsTable"

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Get-Job | Remove-Job -Force
        Write-EventLog @EventEndParams
    }
}