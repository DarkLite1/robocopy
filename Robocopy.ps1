<#
    .SYNOPSIS
        Copy files and folders with Robocopy.exe

    .DESCRIPTION
        Copy files and folders with Robocopy.exe based on its advanced parameters. The parameters will 
        be read from the import file together with the source and the destination folders.
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
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
    }

    $M = "Wait for all $($jobs.count) jobs to be finished"
    Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

    $jobResults = if ($jobs) { $jobs | Wait-Job | Receive-Job }
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

        $robocopyError = $false
        $rows = @()

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
                    $robocopyError = $True 
                }
                default { 
                    $color.Fatal
                    $robocopyError = $True 
                }
            }
            #endregion

            #region Create robocopy log file
            $logParams.Name = $job.Destination + '.log'
            $logFile = New-LogFileNameHC @logParams
            $job.RobocopyOutput | Out-File -LiteralPath $logFile -Encoding utf8
            #endregion

            #region Convert robocopy log file
            $robocopyLogAnalyses = ConvertFrom-RobocopyLogHC -LogFile $LogFile

            $robocopy = @{
                ExitMessage   = ConvertFrom-RobocopyExitCodeHC -ExitCode $job.ExitCode
                ExecutionTime = if (
                    $robocopyLogAnalyses.Times.Total
                ) {
                    $robocopyLogAnalyses.Times.Total 
                }
                else { 'NA' }
                ItemsCopied   = if (
                    $robocopyLogAnalyses.Files.Copied -or 
                    $robocopyLogAnalyses.Dirs.Copied
                ) {
                    [INT]$robocopyLogAnalyses.Files.Copied + 
                    [INT]$robocopyLogAnalyses.Dirs.Copied
                }
                else { 'NA' }
            }
            #endregion

            $rows += @"
<tr bgcolor="$rowColor" style="background:$rowColor;">
<td id="TxtLeft">
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
)<br>
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
)
$(if($job.File){"<br>$($job.File)"})</td>
<td id="TxtLeft">$([String]$robocopy.ItemsCopied + ' (' + $job.ExitCode + ')')</td>
<td id="TxtCentered">$($robocopy.ExecutionTime)</td>
<td id="TxtCentered">$($robocopy.ItemsCopied)</td>
<td id="TxtCentered">$(ConvertTo-HTMLlinkHC -Path $LogFile -Name 'Log')</td>
</tr>
"@
        }

        $HTML = @"
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
</style>
<table id="TxtLeft">
<colgroup><col/><col/><col/><col/><col/></colgroup>
<tr>
<th id="TxtLeft">Robocopy $Switches</th>
<th id="TxtLeft">Launched on</th>
<th id="TxtLeft">Message</th>
<th id="TxtCentered" class="Centered">Total<br>time</th>
<th id="TxtCentered" class="Centered">Copied<br>items</th>
<th id="TxtCentered" class="Centered">Details</th>
</tr>
$rows
</table>
<br>
<style>
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
</style>
<table id="LegendTable">
<tr>
<td bgcolor="$($color.NoCopy)" style="background:$($color.NoCopy);" id="LegendRow">Nothing copied</td>
<td bgcolor="$($color.CopyOk)" style="background:$($color.CopyOk);" id="LegendRow">Copy successful</td>
<td bgcolor="$($color.Mismatch)" style="background:$($color.Mismatch);" id="LegendRow">Clean-up needed</td>
<td bgcolor="$($color.Fatal)" style="background:$($color.Fatal);" id="LegendRow">Fatal error</td>
</tr></table>
"@

        if (($rows.count -eq $FunctionFeed.count) -or !($FunctionFeed.count)) {
            $SuccessfulJobs = "all $($rows.count) jobs."
        }
        else {
            $SuccessfulJobs =
            "only $($rows.count) out of $($FunctionFeed.count) jobs."
        }

        if ($Error) {
            $HTMLErrors = $Error | Get-Unique | 
            ConvertTo-HtmlListHC -Spacing Wide -Header 'Errors detected:'
            $Result = 'FAILURE'
            $Message = Switch ($rows.Count) {
                { $_ -ge 1 } { 
                    "Errors found, we processed $SuccessfulJobs", 
                    $HTMLErrors, $HTML 
                }
                default {
                    "Errors found, we processed $SuccessfulJobs", $HTMLErrors 
                }
            }
            $Priority = 'High'

            $Error | Get-Unique | ForEach-Object {
                Write-EventLog @EventErrorParams -Message $_
            }
        }
        elseif ($robocopyError) {
            $Result = 'FAILURE'
            $Message = "Errors found in the Robocopy log files, we processed $SuccessfulJobs", $HTML
            $Priority = 'High'
            Write-EventLog @EventErrorParams -Message ($env:USERNAME + ' - ' + "$Result`n`n- " + $Message[0])
        }
        else {
            $Result = 'Success'
            $Message = "No errors found, we processed $SuccessfulJobs", $HTML
            $Priority = 'Normal'
            Write-EventLog @EventVerboseParams -Message ($env:USERNAME + ' - ' + "$Result`n`n- " + $Message[0])
        }

        $OutParams = @{
            Path       = $LogFolder
            Name       = "$ScriptName - $Result"
            NamePrefix = 'ScriptStartTime'
            Message    = $Message
        }

        $MailParams = @{
            To        = $MailTo
            Subject   = $Result
            Priority  = $Priority
            Message   = $Message
            LogFolder = $LogFolder
            Header    = $ScriptName
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
        Out-HtmlFileHC @OutParams
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