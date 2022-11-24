#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.Remoting, Toolbox.EventLog

<#
    .SYNOPSIS
        Copy files and folders with Robocopy.exe

    .DESCRIPTION
        Copy files and folders with Robocopy.exe based on its advanced 
        parameters. The parameters will be read from the import file together 
        with the source and the destination folders.

    .PARAMETER MaxConcurrentTasks
        How many robocopy jobs are allowed to run at the same time.

    .PARAMETER SendMail.Header
        The description of the From field displayed in the e-mail client.

    .PARAMETER SendMail.To
        List of e-mail addresses where to send the e-mail too.

    .PARAMETER SendMail.When
        When to send an e-mail. 
        
        Valid options:
        - Always                 : Always send an e-mail
        - Never                  : Never send an e-mail
        - OnlyWhenFilesAreCopied : Only send an e-mail when files are copied

    .PARAMETER RobocopyTasks
        Collection of individual robocopy jobs.

    .PARAMETER RobocopyTasks.Name
        Display a name in the email message instead of the source and 
        destination path.

    .PARAMETER RobocopyTasks.ComputerName
        The computer where to execute the robocopy executable. This allows
        for running robocopy on remote machines. If left blank the job is 
        executed on the current computer.

        To avoid 'Access denied' errors due to the double hop issue it is 
        advised to leave ComputerName blank and use UNC paths in Source and
        Destination.

    .PARAMETER RobocopyTasks.Source
        Specifies the path to the source directory.
        This is the first robocopy argument  known as '<source>'.

    .PARAMETER RobocopyTasks.Destination
        Specifies the path to the destination directory.
        This is the second robocopy argument  known as '<destination>'.

    .PARAMETER RobocopyTasks.File
        Specifies the file or files to be copied. Wildcard characters (* or ?) 
        are supported. If you don't specify this parameter, *.* is used as the 
        default value.
        This is the third robocopy argument known as '<file>'.

    .PARAMETER RobocopyTasks.Switches
        Specifies the options to use with the robocopy command, including copy, 
        file, retry, logging, and job options.
        This is the last robocopy argument known as '<options>'.
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Robocopy\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Function ConvertFrom-RobocopyExitCodeHC {
        <#
        .SYNOPSIS
            Convert exit codes of Robocopy.exe.
    
        .DESCRIPTION
            Convert exit codes of Robocopy.exe to readable formats.
    
        .EXAMPLE
            Robocopy.exe $Source $Target $RobocopySwitches
            ConvertFrom-RobocopyExitCodeHC -ExitCode $LASTEXITCODE
            'COPY'
    
        .NOTES
            $LASTEXITCODE of Robocopy.exe
    
            Hex Bit Value Decimal Value Meaning If Set
            0×10 16 Serious error. Robocopy did not copy any files. This is either 
                 a usage error or an error due to insufficient access privileges on 
                 the source or destination directories.
            0×08 8 Some files or directories could not be copied (copy errors   
                 occurred and the retry limit was exceeded). Check these errors 
                 further.
            0×04 4 Some Mismatched files or directories were detected. Examine the 
                 output log. Housekeeping is probably necessary.
            0×02 2 Some Extra files or directories were detected. Examine the 
                 output log. Some housekeeping may be needed.
            0×01 1 One or more files were copied successfully (that is, new files 
                 have arrived).
            0×00 0 No errors occurred, and no copying was done. The source and 
                 destination directory trees are completely synchronized.
    
            (https://support.microsoft.com/en-us/kb/954404?wa=wsignin1.0)
    
            0	No files were copied. No failure was encountered. No files were 
                mismatched. The files already exist in the destination directory; 
                therefore, the copy operation was skipped.
            1	All files were copied successfully.
            2	There are some additional files in the destination directory that 
                are not present in the source directory. No files were copied.
            3	Some files were copied. Additional files were present. No failure 
                was encountered.
            5	Some files were copied. Some files were mismatched. No failure was 
                encountered.
            6	Additional files and mismatched files exist. No files were copied 
                and no failures were encountered. This means that the files already 
                exist in the destination directory.
            7	Files were copied, a file mismatch was present, and additional  
                files were present.
            8	Several files did not copy.
            
            * Note Any value greater than 8 indicates that there was at least one 
            failure during the copy operation.
            #>
    
        Param (
            [int]$ExitCode
        )
    
        Process {
            Switch ($ExitCode) {
                0 { $Message = 'NO CHANGE'; break }
                1 { $Message = 'COPY'; break }
                2 { $Message = 'EXTRA'; break }
                3 { $Message = 'EXTRA + COPY'; break }
                4 { $Message = 'MISMATCH'; break }
                5 { $Message = 'MISMATCH + COPY'; break }
                6 { $Message = 'MISMATCH + EXTRA'; break }
                7 { $Message = 'MISMATCH + EXTRA + COPY'; break }
                8 { $Message = 'FAIL'; break }
                9 { $Message = 'FAIL + COPY'; break }
                10 { $Message = 'FAIL + EXTRA'; break }
                11 { $Message = 'FAIL + EXTRA + COPY'; break }
                12 { $Message = 'FAIL + MISMATCH'; break }
                13 { $Message = 'FAIL + MISMATCH + COPY'; break }
                14 { $Message = 'FAIL + MISMATCH + EXTRA'; break }
                15 { $Message = 'FAIL + MISMATCH + EXTRA + COPY'; break }
                16 { $Message = 'FATAL ERROR'; break }
                default { 'UNKNOWN' }
            }
            return $Message
        }
    }
    Function ConvertFrom-RobocopyLogHC {
        <#
            .SYNOPSIS
                Create a PSCustomObject from a Robocopy log file.
    
            .DESCRIPTION
                Parses Robocopy logs into a collection of objects summarizing each 
                Robocopy operation.
    
            .EXAMPLE
                ConvertFrom-RobocopyLogHC 'C:\robocopy.log'
                Source      : \\contoso.net\folder1\
                Destination : \\contoso.net\folder2\
                Dirs        : @{
                    Total=2; Copied=0; Skipped=2; 
                    Mismatch=0; FAILED=0; Extras=0
                }
                Files       : @{
                    Total=203; Copied=0; Skipped=203; 
                    Mismatch=0; FAILED=0; Extras=0
                }
                Times       : @{
                    Total=0:00:00; Copied=0:00:00; 
                    FAILED=0:00:00; Extras=0:00:00
                }
    #>
    
        Param (
            [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
            [ValidateScript( { Test-Path $_ -PathType Leaf })]
            [String]$LogFile
        )
    
        Process {
            $Header = Get-Content $LogFile | Select-Object -First 12
            $Footer = Get-Content $LogFile | Select-Object -Last 9
    
            $Header | ForEach-Object {
                if ($_ -like "*Source :*") {
                    $Source = (($_.Split(':', 2))[1]).trim()
                }
                if ($_ -like "*Dest :*") {
                    $Destination = (($_.Split(':', 2))[1]).trim()
                }
                # in case of robo error log
                if ($_ -like "*Source -*") {
                    $Source = (($_.Split('-', 2))[1]).trim()
                }
                if ($_ -like "*Dest -*") {
                    $Destination = (($_.Split('-', 2))[1]).trim()
                }
            }
    
            $Footer | ForEach-Object {
                if ($_ -like "*Dirs :*") {
                    $Array = (($_.Split(':')[1]).trim()) -split '\s+'
                    $Dirs = [PSCustomObject][Ordered]@{
                        Total    = $Array[0]
                        Copied   = $Array[1]
                        Skipped  = $Array[2]
                        Mismatch = $Array[3]
                        FAILED   = $Array[4]
                        Extras   = $Array[5]
                    }
                }
                if ($_ -like "*Files :*") {
                    $Array = ($_.Split(':')[1]).trim() -split '\s+'
                    $Files = [PSCustomObject][Ordered]@{
                        Total    = $Array[0]
                        Copied   = $Array[1]
                        Skipped  = $Array[2]
                        Mismatch = $Array[3]
                        FAILED   = $Array[4]
                        Extras   = $Array[5]
                    }
                }
                if ($_ -like "*Times :*") {
                    $Array = ($_.Split(':', 2)[1]).trim() -split '\s+'
                    $Times = [PSCustomObject][Ordered]@{
                        Total  = $Array[0]
                        Copied = $Array[1]
                        FAILED = $Array[2]
                        Extras = $Array[3]
                    }
                }
            }
    
            $Obj = [PSCustomObject][Ordered]@{
                'Source'      = $Source
                'Destination' = $Destination
                'Dirs'        = $Dirs
                'Files'       = $Files
                'Times'       = $Times
            }
            Write-Output $Obj
        }
    }

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
            
            $global:LASTEXITCODE = 0 # required to get the correct exit code
            
            $expression = [String]::Format(
                'ROBOCOPY "{0}" "{1}" {2} {3}', 
                $Source, $Destination, $File, $Switches
            )
            $result.RobocopyOutput = Invoke-Expression $expression
            $result.ExitCode = $LASTEXITCODE
        }
        Catch {
            $result.Error = $_
        }
        Finally {
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
        if (-not ($mailTo = $file.SendMail.To)) {
            throw "Input file '$ImportFile': No 'SendMail.To' addresses found."
        }
        if (-not ($mailWhen = $file.SendMail.When)) {
            throw "Input file '$ImportFile': No 'SendMail.When' found, valid options are: Always, Never or OnlyWhenFilesAreCopied."
        }
        $mailHeader = $file.SendMail.Header
        
        if (-not ($RobocopyTasks = $file.RobocopyTasks)) {
            throw "Input file '$ImportFile': No 'RobocopyTasks' found."
        }
        foreach ($task in $RobocopyTasks) {
            #region Mandatory parameters
            if (-not $task.Source) {
                throw "Input file '$ImportFile': No 'Source' found in one of the 'RobocopyTasks'."
            }
            if (-not $task.Destination) {
                throw "Input file '$ImportFile': No 'Destination' found for source '$($task.Source)'."
            }
            if (-not $task.Switches) {
                throw "Input file '$ImportFile': No 'Switches' found for source '$($task.Source)'."
            }
            #endregion

            #region Avoid double hop issue
            if (
                ($task.ComputerName) -and
                (
                    ($task.Source -Match '^\\\\') -or 
                    ($task.Destination -Match '^\\\\')
                )
            ) {
                throw "Input file '$ImportFile' ComputerName '$($task.ComputerName)', Source '$($task.Source)', Destination '$($task.Destination)': When ComputerName is used only local paths are allowed. This to avoid the double hop issue."
            }
            #endregion

            #region Avoid mix of local paths with UNC paths
            if (
                (-not $task.ComputerName) -and
                (
                    ($task.Source -notMatch '^\\\\') -or 
                    ($task.Destination -notMatch '^\\\\')
                )
            ) {
                throw "Input file '$ImportFile' Source '$($task.Source)', Destination '$($task.Destination)': When ComputerName is not used only UNC paths are allowed."
            }
            #endregion
        }

        if ($file.PSObject.Properties.Name -notContains 'MaxConcurrentJobs') {
            throw "Input file '$ImportFile': Property 'MaxConcurrentJobs' not found."
        }
        if (-not ($file.MaxConcurrentJobs -is [int])) {
            throw "Input file '$ImportFile': Property 'MaxConcurrentJobs' needs to be a number, the value '$($file.MaxConcurrentJobs)' is not supported."
        }

        $MaxConcurrentJobs = [int]$file.MaxConcurrentJobs
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
            totalFilesCopied    = 0
            robocopyBadExitCode = 0
            robocopyJobError    = 0
            systemError         = 0
        }
        
        $htmlTableRows = @() 

        Foreach ($job in $jobResults) {
            $M = "Job result: Name '$($job.Name), ComputerName '$($job.ComputerName)', Source '$($job.Source)', Destination '$($job.Destination)', File '$($job.File)', Switches '$($job.Switches)', ExitCode '$($job.ExitCode)', Error '$($job.Error)'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

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

            $filesCopiedCount = [INT]$robocopyLogAnalyses.Files.Copied
            $counter.totalFilesCopied += $filesCopiedCount

            $robocopy = @{
                ExitMessage   = ConvertFrom-RobocopyExitCodeHC -ExitCode $job.ExitCode
                ExecutionTime = if (
                    $robocopyLogAnalyses.Times.Total
                ) {
                    $robocopyLogAnalyses.Times.Total 
                }
                else { 'NA' }
                FilesCopied   = $filesCopiedCount 
            }
            #endregion

            #region Create HTML table rows
            $htmlTableRows += @"
<tr bgcolor="$rowColor" style="background:$rowColor;">
    <td id="TxtLeft">{0}<br>{1}{2}{3}</td>
    <td id="TxtLeft">$($robocopy.ExitMessage + ' (' + $job.ExitCode + ')')</td>
    <td id="TxtCentered">$($robocopy.ExecutionTime)</td>
    <td id="TxtCentered">$($robocopy.FilesCopied)</td>
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
                <th id="TxtCentered" class="Centered">Files<br>copied</th>
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
        $logParams.Name = "$ScriptName - Mail.html"

        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Priority  = 'Normal' 
            Subject   = '{0} job{1}, {2} file{3} copied' -f 
            $RobocopyTasks.Count, 
            $(if ($RobocopyTasks.Count -ne 1) { 's' }), 
            $counter.totalFilesCopied,
            $(if ($counter.totalFilesCopied -ne 1) { 's' })
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
            $allErrorCount, 
            $(if ($allErrorCount -ne 1) { 's' }), 
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
            </table><br>
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