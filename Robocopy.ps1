#Requires -Version 7

<#
    .SYNOPSIS
        Copy files and folders with Robocopy.exe

    .DESCRIPTION
        Copy/move/mirror files and folders with Robocopy.exe based on its
        advanced parameters. The parameters are read from the import file.

        Send an e-mail to the user when needed, but always send an e-mail to
        the admin on errors.

    .PARAMETER MaxConcurrentTasks
        How many robocopy jobs are allowed to run at the same time.

    .PARAMETER SendMail.Header
        The description of the From field displayed in the e-mail client.

    .PARAMETER SendMail.To
        List of e-mail addresses where to send the e-mail too.

    .PARAMETER SendMail.When
        When to send an e-mail to the user.

        Valid options:
        - Never               : Never send an e-mail
        - Always              : Always send an e-mail
        - OnlyOnError         : Send no e-mail except when errors are found
        - OnlyOnErrorOrAction : Only send an e-mail when files are copied or
                                errors are found

        The script admin will always receive an e-mail.

    .PARAMETER Tasks
        Collection of individual robocopy jobs.

    .PARAMETER Tasks.Name
        Display a name in the email message instead of the source and
        destination path.

    .PARAMETER Tasks.ComputerName
        The computer where to execute the robocopy executable. This allows
        for running robocopy on remote machines. If left blank the job is
        executed on the current computer.

        To avoid 'Access denied' errors due to the double hop issue it is
        advised to leave ComputerName blank and use UNC paths in Source and
        Destination.

    .PARAMETER Tasks.Robocopy.InputFile
        Specifies the path to the input file that robocopy can use.

        - 1) Create a robocopy input file

            Careful, because robocopy will be executed with the provided
            arguments while creating the job file.

            Example:
            Robocopy.exe /MOV C:\SourceFolder C:\DestinationFolder /R:5 /W:30 /ZB /XF ExcludedFile.txt /SAVE:C:\robocopyConfig

        - 2) Edit the input file to your requirements

            Manually in a text editor or by using code:

            New-TextFileHC -InputPath 'C:\robocopyConfig.RCJ' -ReplaceLine '		ExcludedFile.txt' -NewLine @(
                '		ExcludeFile1.txt',
                '		ExcludeFile2.txt',
                '		ExcludeFile3.txt'
            ) -NewFilePath 'C:\robocopyConfig.RCJ' -Overwrite

        - 3) Set Tasks.Robocopy.InputFile

            Tasks.Robocopy.InputFile = C:\robocopyConfig
            (runs: Robocopy.exe /job:robocopyConfig)

    .PARAMETER Tasks.Robocopy.Arguments.Source
        Specifies the path to the source directory.
        This is the first robocopy argument  known as '<source>'.

    .PARAMETER Tasks.Robocopy.Arguments.Destination
        Specifies the path to the destination directory.
        This is the second robocopy argument  known as '<destination>'.

    .PARAMETER Tasks.Robocopy.Arguments.File
        Specifies the file or files to be copied. Wildcard characters (* or ?)
        are supported. If you don't specify this parameter, *.* is used as the
        default value.
        This is the third robocopy argument known as '<file>'.

    .PARAMETER Tasks.Robocopy.Arguments.Switches
        Specifies the options to use with the robocopy command, including copy,
        file, retry, logging, and job options.
        This is the last robocopy argument known as '<options>'.
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [String]$ConfigurationJsonFile
)

begin {
    $ErrorActionPreference = 'stop'

    $eventLogData = [System.Collections.Generic.List[PSObject]]::new()
    $systemErrors = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date

    try {
        Function Convert-RobocopyExitCodeToStringHC {
            <#
            .SYNOPSIS
                Convert exit codes of Robocopy.exe to strings.

            .EXAMPLE
                Robocopy.exe $Source $Target $RobocopySwitches

                Convert-RobocopyExitCodeToStringHC -ExitCode $LASTEXITCODE

                Returns: 'COPY'
            #>

            Param (
                [int]$ExitCode
            )

            switch ($ExitCode) {
                0 { return 'NO CHANGE' }
                1 { return 'COPY' }
                2 { return 'EXTRA' }
                3 { return 'EXTRA + COPY' }
                4 { return 'MISMATCH' }
                5 { return 'MISMATCH + COPY' }
                6 { return 'MISMATCH + EXTRA' }
                7 { return 'MISMATCH + EXTRA + COPY' }
                8 { return 'FAIL' }
                9 { return 'FAIL + COPY' }
                10 { return 'FAIL + EXTRA' }
                11 { return 'FAIL + EXTRA + COPY' }
                12 { return 'FAIL + MISMATCH' }
                13 { return 'FAIL + MISMATCH + COPY' }
                14 { return 'FAIL + MISMATCH + EXTRA' }
                15 { return 'FAIL + MISMATCH + EXTRA + COPY' }
                16 { return 'FATAL ERROR' }
                default { return 'UNKNOWN ROBOCOPY EXIT CODE' }
            }
        }

        Function Convert-RobocopyLogToObjectHC {
            Param (
                [String[]]$LogContent
            )

            $result = [ordered]@{
                'Source'      = ''
                'Destination' = ''
                'Dirs'        = [PSCustomObject]@{
                    Total    = 0
                    Copied   = 0
                    Skipped  = 0
                    Mismatch = 0
                    FAILED   = 0
                    Extras   = 0
                }
                'Files'       = [PSCustomObject]@{
                    Total    = 0
                    Copied   = 0
                    Skipped  = 0
                    Mismatch = 0
                    FAILED   = 0
                    Extras   = 0
                }
                'Times'       = [PSCustomObject]@{
                    Total  = ''
                    Copied = ''
                    FAILED = ''
                    Extras = ''
                }
            }

            $LogContent | ForEach-Object {
                if ($_ -match '^\s*(?:Source|Dest)\s*[:=-]\s*(.*)') {
                    $path = $Matches[1].Trim()

                    if ($_ -match 'Source') {
                        $result.Source = $path
                    }
                    else {
                        $result.Destination = $path
                    }
                }
                elseif (
                    $_ -match '^\s*(Dirs|Files)\s*:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s*$'
                ) {
                    $key = $Matches[1]
                    $values = ($Matches[2..$Matches.Count] -join ' ').Split(
                        ' ', [System.StringSplitOptions]::RemoveEmptyEntries
                    )

                    switch ($key) {
                        'Dirs' {
                            $result.Dirs.Total = [int]$values[0]
                            $result.Dirs.Copied = [int]$values[1]
                            $result.Dirs.Skipped = [int]$values[2]
                            $result.Dirs.Mismatch = [int]$values[3]
                            $result.Dirs.FAILED = [int]$values[4]
                            $result.Dirs.Extras = [int]$values[5]
                        }
                        'Files' {
                            $result.Files.Total = [int]$values[0]
                            $result.Files.Copied = [int]$values[1]
                            $result.Files.Skipped = [int]$values[2]
                            $result.Files.Mismatch = [int]$values[3]
                            $result.Files.FAILED = [int]$values[4]
                            $result.Files.Extras = [int]$values[5]
                        }
                    }
                }
                elseif (
                    $_ -match '^\s*Times\s*:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s*$'
                ) {
                    $values = $Matches[1..4]
                    $result.Times.Total = $values[0]
                    $result.Times.Copied = $values[1]
                    $result.Times.FAILED = $values[2]
                    $result.Times.Extras = $values[3]
                }
            }

            [PSCustomObject]$result
        }

        function Get-StringValueHC {
            <#
        .SYNOPSIS
            Retrieve a string from the environment variables or a regular string.

        .DESCRIPTION
            This function checks the 'Name' property. If the value starts with
            'ENV:', it attempts to retrieve the string value from the specified
            environment variable. Otherwise, it returns the value directly.

        .PARAMETER Name
            Either a string starting with 'ENV:'; a plain text string or NULL.

        .EXAMPLE
            Get-StringValueHC -Name 'ENV:passwordVariable'

            # Output: the environment variable value of $ENV:passwordVariable
            # or an error when the variable does not exist

        .EXAMPLE
            Get-StringValueHC -Name 'mySecretPassword'

            # Output: mySecretPassword

        .EXAMPLE
            Get-StringValueHC -Name ''

            # Output: NULL
        #>
            param (
                [String]$Name
            )

            if (-not $Name) {
                return $null
            }
            elseif (
                $Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)
            ) {
                $envVariableName = $Name.Substring(4).Trim()
                $envStringValue = Get-Item -Path "Env:\$envVariableName" -EA Ignore
                if ($envStringValue) {
                    return $envStringValue.Value
                }
                else {
                    throw "Environment variable '$envVariableName' not found."
                }
            }
            else {
                return $Name
            }
        }

        $eventLogData.Add(
            [PSCustomObject]@{
                Message   = 'Script started'
                DateTime  = $scriptStartTime
                EntryType = 'Information'
                EventID   = '100'
            }
        )

        #region Import .json file
        Write-Verbose "Import .json file '$ConfigurationJsonFile'"

        $jsonFileItem = Get-Item -LiteralPath $ConfigurationJsonFile -ErrorAction Stop

        $jsonFileContent = Get-Content $jsonFileItem -Raw -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        @(
            'MaxConcurrentTasks', 'Tasks'
        ).where(
            { -not $jsonFileContent.$_ }
        ).foreach(
            { throw "Property '$_' not found" }
        )

        #region Test integer value
        try {
            [int]$MaxConcurrentTasks = $jsonFileContent.MaxConcurrentTasks
        }
        catch {
            throw "Property 'MaxConcurrentTasks' needs to be a number, the value '$($jsonFileContent.MaxConcurrentTasks)' is not supported."
        }
        #endregion

        $Tasks = $jsonFileContent.Tasks

        foreach ($task in $Tasks) {
            if ($task.Robocopy.Arguments) {
                if ($task.Robocopy.InputFile) {
                    throw "Property 'Tasks.Robocopy.Arguments' and 'Tasks.Robocopy.InputFile' cannot be used at the same time"
                }

                #region Mandatory parameters
                @(
                    'Source', 'Destination', 'Switches'
                ).where(
                    { -not $jsonFileContent.Tasks.Robocopy.Arguments.$_ }
                ).foreach(
                    { throw "Property 'Tasks.Robocopy.Arguments.$_' not found" }
                )
                #endregion

                #region Avoid double hop issue
                if (
                    ($task.ComputerName) -and
                    (
                        ($task.Robocopy.Arguments.Source -Match '^\\\\') -or
                        ($task.Robocopy.Arguments.Destination -Match '^\\\\')
                    )
                ) {
                    throw "ComputerName '$($task.ComputerName)', Source '$($task.Robocopy.Arguments.Source)', Destination '$($task.Robocopy.Arguments.Destination)': When ComputerName is used only local paths are allowed. This to avoid the double hop issue."
                }
                #endregion

                #region Avoid mix of local paths with UNC paths
                if (
                    (-not $task.ComputerName) -and
                    (
                        ($task.Robocopy.Arguments.Source -notMatch '^\\\\') -or
                        ($task.Robocopy.Arguments.Destination -notMatch '^\\\\')
                    )
                ) {
                    throw "Source '$($task.Robocopy.Arguments.Source)', Destination '$($task.Robocopy.Arguments.Destination)': When ComputerName is not used only UNC paths are allowed."
                }
                #endregion
            }
            elseif ($task.Robocopy.InputFile) {
                if (
                    -not (
                        Test-Path -Path $task.Robocopy.InputFile -PathType Leaf
                    )
                ) {
                    throw "Property 'Tasks.Robocopy.InputFile' path '$($task.Robocopy.InputFile)' not found"
                }
            }
            else {
                throw "Property 'Tasks.Robocopy.Arguments' or 'Tasks.Robocopy.InputFile' not found"
            }

        }
        #endregion

        #region Convert .json file
        Write-Verbose 'Convert .json file'

        #region Set PSSessionConfiguration
        $PSSessionConfiguration = $jsonFileContent.PSSessionConfiguration

        if (-not $PSSessionConfiguration) {
            $PSSessionConfiguration = 'PowerShell.7'
        }
        #endregion

        foreach ($task in $Tasks) {
            #region Set ComputerName if there is none
            if (
                (-not $task.ComputerName) -or
                ($task.ComputerName -eq 'localhost') -or
                ($task.ComputerName -eq "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
            ) {
                $task.ComputerName = $env:COMPUTERNAME
            }
            #endregion

            #region Add properties
            $task | Add-Member -NotePropertyMembers @{
                Job = @{
                    Results = @()
                    Error   = $null
                }
            }
            #endregion
        }
        #endregion
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Input file '$ConfigurationJsonFile': $_"
            }
        )

        Write-Warning $systemErrors[-1].Message

        return
    }
}

process {
    if ($systemErrors) { return }

    try {
        $scriptBlock = {
            Try {
                $task = $_

                #region Declare variables for parallel execution
                if (-not $MaxConcurrentTasks) {
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                    $eventLogData = $using:eventLogData
                }
                #endregion

                if ($task.Robocopy.InputFile) {
                    $invokeParams = @{
                        ArgumentList = $task.Robocopy.InputFile,
                        $task.TaskName, $task.ComputerName
                        ScriptBlock  = {
                            Param (
                                [Parameter(Mandatory)]
                                [String]$InputFile,
                                [String]$Name,
                                [String]$ComputerName
                            )

                            Try {
                                $result = [PSCustomObject]@{
                                    Name           = $Name
                                    ComputerName   = $ComputerName
                                    InputFile      = $InputFile
                                    Source         = $null
                                    Destination    = $null
                                    File           = $null
                                    Switches       = $null
                                    RobocopyOutput = $null
                                    ExitCode       = $null
                                    Error          = $null
                                }

                                #region Copy input file to temp file
                                # only local paths are supported by /job
                                try {
                                    $joinParams = @{
                                        Path      = $env:TEMP
                                        ChildPath = ([System.IO.Path]::GetFileName($InputFile))
                                    }
                                    $tempJobFile = Join-Path @joinParams

                                    Copy-Item -Path $InputFile -Destination $tempJobFile -Force
                                }
                                catch {
                                    throw "Failed to copy job file '$InputFile' to temp file on '$($env:COMPUTERNAME)': $_"
                                }
                                #endregion

                                $global:LASTEXITCODE = 0

                                $expression = [String]::Format(
                                    "ROBOCOPY /job:`"$tempJobFile`""
                                )
                                $result.RobocopyOutput = Invoke-Expression $expression
                                $result.ExitCode = $LASTEXITCODE
                            }
                            Catch {
                                $result.Error = $_
                            }
                            Finally {
                                Remove-Item $tempJobFile -Force -ErrorAction Ignore

                                $result
                            }
                        }
                    }

                    #region Verbose
                    $M = "Start job on '{0}' with TaskName '{1}' InputFile '{2}'" -f $task.ComputerName,
                    $invokeParams.ArgumentList[1],
                    $invokeParams.ArgumentList[0]

                    Write-Verbose $M

                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = $M
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '2'
                        }
                    )
                    #endregion
                }
                else {
                    $invokeParams = @{
                        ArgumentList = $task.Robocopy.Arguments.Source,
                        $task.Robocopy.Arguments.Destination,
                        $task.Robocopy.Arguments.Switches,
                        $task.Robocopy.Arguments.File,
                        $task.TaskName, $task.ComputerName
                        ScriptBlock  = {
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
                                    InputFile      = $null
                                    Source         = $Source
                                    Destination    = $Destination
                                    File           = $File
                                    Switches       = $Switches
                                    RobocopyOutput = $null
                                    ExitCode       = $null
                                    Error          = $null
                                }

                                $global:LASTEXITCODE = 0

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
                    }

                    #region Verbose
                    $M = "Start job on '{0}' with Source '{1}' Destination '{2}' Switches '{3}' File '{4}' TaskName '{5}'" -f $task.ComputerName,
                    $invokeParams.ArgumentList[0],
                    $invokeParams.ArgumentList[1],
                    $invokeParams.ArgumentList[2],
                    $invokeParams.ArgumentList[3],
                    $invokeParams.ArgumentList[4]

                    Write-Verbose $M

                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = $M
                            DateTime  = Get-Date
                            EntryType = 'Information'
                            EventID   = '2'
                        }
                    )
                    #endregion
                }

                #region Start job
                $computerName = $task.ComputerName

                $task.Job.Results += if (
                    $computerName -eq $ENV:COMPUTERNAME
                ) {
                    $params = $invokeParams.ArgumentList
                    & $invokeParams.ScriptBlock @params
                }
                else {
                    $invokeParams += @{
                        ConfigurationName   = $PSSessionConfiguration
                        ComputerName        = $computerName
                        EnableNetworkAccess = $true
                        ErrorAction         = 'Stop'
                    }
                    Invoke-Command @invokeParams
                }
                #endregion
            }
            catch {
                $task.Job.Error = $_
                $Error.RemoveAt(0)
            }
        }

        #region Run code serial or parallel
        $foreachParams = if ($MaxConcurrentTasks -eq 1) {
            @{
                Process = $scriptBlock
            }
        }
        else {
            @{
                Parallel      = $scriptBlock
                ThrottleLimit = $MaxConcurrentTasks
            }
        }
        #endregion

        $Tasks | ForEach-Object @foreachParams

        Write-Verbose 'All tasks finished'
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = $_
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
}

end {
    function ConvertTo-HtmlListHC {
        <#
        .SYNOPSIS
            Creates an unordered HTML list.

        .PARAMETER Message
            The items in the list.

        .PARAMETER Header
            Add a header '<h3>My list title</h3>' above the unordered list.

        .PARAMETER FootNote
            Add a small text at the bottom of the unordered list in a smaller
            font and italic. This is convenient for adding a small explanation
            of the items or a legend.

        .EXAMPLE
            $params = [ordered]@{
                Message = @('Item 1', 'Item 2')
            }
            ConvertTo-HtmlListHC @params

            Create the following HTML code:
            '<ul>
                <li style="margin: 10px 0;">Item 1</li>
                <li style="margin: 10px 0;">Item 2</li>
            </ul>'
        #>
        Param (
            [parameter(Mandatory, ValueFromPipeline)]
            [String[]]$Message,
            [String]$Header,
            [String]$FootNote
        )

        begin {
            $allItems = [System.Collections.ArrayList]::new()
        }

        process {
            $null = $allItems.AddRange($Message)
        }

        end {
            @"
$($Header ? "<h3>$Header</h3>" : '')
<ul>
    $(
        $allItems |
        ForEach-Object { "<li style=`"margin: 10px 0;`">$_</li>" }
    )
</ul>
$($FootNote ? "<i><font size=`"2`">* $FootNote</font></i>" : '')
"@
        }
    }

    function Get-LogFolderHC {
        <#
        .SYNOPSIS
            Ensures that a specified path exists, creating it if it doesn't.
            Supports absolute paths and paths relative to $PSScriptRoot. Returns
            the full path of the folder.

            .DESCRIPTION
            This function takes a path as input and checks if it exists. if
            the path does not exist, it attempts to create the folder. It
            handles both absolute paths and paths relative to the location of
            the currently running script ($PSScriptRoot).

            .PARAMETER Path
            The path to ensure exists. This can be an absolute path (ex.
                C:\MyFolder\SubFolder) or a path relative to the script's
            directory (ex. Data\Logs).

        .EXAMPLE
            Get-LogFolderHC -Path 'C:\MyData\Output'
            # Ensures the directory 'C:\MyData\Output' exists.

        .EXAMPLE
            Get-LogFolderHC -Path 'Logs\Archive'
            # If the script is in 'C:\Scripts', this ensures 'C:\Scripts\Logs\Archive' exists.
        #>

        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if ($Path -match '^[a-zA-Z]:\\' -or $Path -match '^\\') {
            $fullPath = $Path
        }
        else {
            $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $Path
        }

        if (-not (Test-Path -Path $fullPath -PathType Container)) {
            try {
                Write-Verbose "Create log folder '$fullPath'"
                $null = New-Item -Path $fullPath -ItemType Directory -Force
            }
            catch {
                throw "Failed creating log folder '$fullPath': $_"
            }
        }

        (Resolve-Path $fullPath).ProviderPath
        # $fullPath
    }

    function Get-ValidFileNameHC {
        param (
            [string]$Path
        )

        $invalidChars = '[<>:"/\\|?*]'
        $validFileName = $Path -replace $invalidChars, '_'

        return $validFileName
    }

    function Out-LogFileHC {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [PSCustomObject[]]$DataToExport,
            [Parameter(Mandatory)]
            [String]$PartialPath,
            [Parameter(Mandatory)]
            [String[]]$FileExtensions,
            [hashtable]$ExcelFile = @{
                SheetName = 'Overview'
                TableName = 'Overview'
                CellStyle = $null
            },
            [Switch]$Append
        )

        $allLogFilePaths = @()

        foreach (
            $fileExtension in
            $FileExtensions | Sort-Object -Unique
        ) {
            try {
                $logFilePath = "$PartialPath{0}" -f $fileExtension

                $M = "Export {0} object{1} to '$logFilePath'" -f
                $DataToExport.Count,
                $(if ($DataToExport.Count -ne 1) { 's' })
                Write-Verbose $M

                switch ($fileExtension) {
                    '.csv' {
                        $params = @{
                            LiteralPath       = $logFilePath
                            Append            = $Append
                            Delimiter         = ';'
                            NoTypeInformation = $true
                        }
                        $DataToExport | Export-Csv @params

                        break
                    }
                    '.json' {
                        #region Convert error object to error message string
                        $convertedDataToExport = foreach (
                            $exportObject in
                            $DataToExport
                        ) {
                            foreach ($property in $exportObject.PSObject.Properties) {
                                $name = $property.Name
                                $value = $property.Value
                                if (
                                    $value -is [System.Management.Automation.ErrorRecord]
                                ) {
                                    if (
                                        $value.Exception -and $value.Exception.Message
                                    ) {
                                        $exportObject.$name = $value.Exception.Message
                                    }
                                    else {
                                        $exportObject.$name = $value.ToString()
                                    }
                                }
                            }
                            $exportObject
                        }
                        #endregion

                        if (
                            $Append -and
                            (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                        ) {
                            $params = @{
                                LiteralPath = $logFilePath
                                Raw         = $true
                                Encoding    = 'UTF8'
                            }
                            $jsonFileContent = Get-Content @params | ConvertFrom-Json

                            $convertedDataToExport = [array]$convertedDataToExport + [array]$jsonFileContent
                        }

                        $convertedDataToExport |
                        ConvertTo-Json -Depth 7 |
                        Out-File -LiteralPath $logFilePath

                        break
                    }
                    '.txt' {
                        $params = @{
                            LiteralPath = $logFilePath
                            Append      = $Append
                        }

                        $DataToExport | Format-List -Property * -Force |
                        Out-File @params

                        break
                    }
                    '.xlsx' {
                        if (
                            (-not $Append) -and
                            (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                        ) {
                            $logFilePath | Remove-Item
                        }

                        $excelParams = @{
                            Path          = $logFilePath
                            Append        = $true
                            AutoNameRange = $true
                            AutoSize      = $true
                            FreezeTopRow  = $true
                            WorksheetName = $ExcelFile.SheetName
                            TableName     = $ExcelFile.TableName
                            Verbose       = $false
                        }
                        if ($ExcelFile.CellStyle) {
                            $excelParams.CellStyleSB = $ExcelFile.CellStyle
                        }
                        $DataToExport | Export-Excel @excelParams

                        break
                    }
                    default {
                        throw "Log file extension '$_' not supported. Supported values are '.csv', '.json', '.txt' or '.xlsx'."
                    }
                }

                $allLogFilePaths += $logFilePath
            }
            catch {
                Write-Warning "Failed creating log file '$logFilePath': $_"
            }
        }

        $allLogFilePaths
    }

    function Send-MailKitMessageHC {
        <#
            .SYNOPSIS
                Send an email using MailKit and MimeKit assemblies.

            .DESCRIPTION
                This function sends an email using the MailKit and MimeKit
                assemblies. It requires the assemblies to be installed before
                calling the function:

                $params = @{
                    Source           = 'https://www.nuget.org/api/v2'
                    SkipDependencies = $true
                    Scope            = 'AllUsers'
                }
                Install-Package @params -Name 'MailKit'
                Install-Package @params -Name 'MimeKit'

            .PARAMETER MailKitAssemblyPath
                The path to the MailKit assembly.

            .PARAMETER MimeKitAssemblyPath
                The path to the MimeKit assembly.

            .PARAMETER SmtpServerName
                The name of the SMTP server.

            .PARAMETER SmtpPort
                The port of the SMTP server.

            .PARAMETER SmtpConnectionType
                The connection type for the SMTP server.

                Valid values are:
                - 'None'
                - 'Auto'
                - 'SslOnConnect'
                - 'StartTlsWhenAvailable'
                - 'StartTls'

            .PARAMETER Credential
                The credential object containing the username and password.

            .PARAMETER From
                The sender's email address.

            .PARAMETER FromDisplayName
            The display name to show for the sender.

            Email clients may display this differently. It is most likely
            to be shown if the sender's email address is not recognized
                (e.g., not in the address book).

            .PARAMETER To
                The recipient's email address.

            .PARAMETER Body
            The body of the email, HTML is supported.

            .PARAMETER Subject
            The subject of the email.

            .PARAMETER Attachments
            An array of file paths to attach to the email.

            .PARAMETER Priority
            The email priority.

            Valid values are:
            - 'Low'
            - 'Normal'
            - 'High'

            .EXAMPLE
            # Send an email with StartTls and credential

            $SmtpUserName = 'smtpUser'
            $SmtpPassword = 'smtpPassword'

            $securePassword = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($SmtpUserName, $securePassword)

            $params = @{
                SmtpServerName = 'SMT_SERVER@example.com'
                SmtpPort = 587
                SmtpConnectionType = 'StartTls'
                Credential = $credential
                from = 'm@example.com'
                To = '007@example.com'
                Body = '<p>Mission details in attachment</p>'
                Subject = 'For your eyes only'
                Priority = 'High'
                Attachments = @('c:\Mission.ppt', 'c:\ID.pdf')
                MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
            }

            Send-MailKitMessageHC @params

            .EXAMPLE
            # Send an email without authentication

            $params = @{
                SmtpServerName      = 'SMT_SERVER@example.com'
                SmtpPort            = 25
                From                = 'hacker@example.com'
                FromDisplayName     = 'White hat hacker'
                Bcc                 = @('james@example.com', 'mike@example.com')
                Body                = '<h1>You have been hacked</h1>'
                Subject             = 'Oops'
                MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
            }

            Send-MailKitMessageHC @params
            #>

        [CmdletBinding()]
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
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
            [parameter(Mandatory)]
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [string]$FromDisplayName,
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

        begin {
            function Test-IsAssemblyLoaded {
                param (
                    [String]$Name
                )
                foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                    if ($assembly.FullName -like "$Name, Version=*") {
                        return $true
                    }
                }
                return $false
            }

            function Add-Attachments {
                param (
                    [string[]]$Attachments,
                    [MimeKit.Multipart]$BodyMultiPart
                )

                $attachmentList = New-Object System.Collections.ArrayList($null)

                foreach (
                    $attachmentPath in
                    $Attachments | Sort-Object -Unique
                ) {
                    try {
                        #region Test if file exists
                        try {
                            $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                            if ($attachmentItem.PSIsContainer) {
                                Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                                continue
                            }
                        }
                        catch {
                            Write-Warning "Attachment '$attachmentPath' not found"
                            continue
                        }
                        #endregion

                        $totalSizeAttachments += $attachmentItem.Length

                        $null = $attachmentList.Add($attachmentItem)

                        #region Check size of attachments
                        if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                            $M = "The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details." -f
                            ([math]::Round(($MaxAttachmentSize / 1MB))),
                            ([math]::Round(($totalSizeAttachments / 1MB), 2))

                            Write-Warning $M

                            return [PSCustomObject]@{
                                AttachmentLimitExceededMessage = $M
                            }
                        }
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentPath': $_"
                    }
                }
                #endregion

                foreach (
                    $attachmentItem in
                    $attachmentList
                ) {
                    try {
                        Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                        $attachment = New-Object MimeKit.MimePart

                        #region Create a MemoryStream to hold the file content
                        $memoryStream = New-Object System.IO.MemoryStream

                        try {
                            $fileStream = [System.IO.File]::OpenRead($attachmentItem.FullName)
                            $fileStream.CopyTo($memoryStream)
                        }
                        finally {
                            if ($fileStream) {
                                $fileStream.Dispose()
                            }
                        }

                        $memoryStream.Position = 0
                        #endregion

                        $attachment.Content = New-Object MimeKit.MimeContent($memoryStream)

                        $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                        $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                        $attachment.FileName = $attachmentItem.Name

                        $bodyMultiPart.Add($attachment)
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentItem': $_"
                    }
                }
            }

            try {
                #region Test To or Bcc required
                if (-not ($To -or $Bcc)) {
                    throw "Either 'To' to 'Bcc' is required for sending emails"
                }
                #endregion

                #region Test To
                foreach ($email in $To) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "To email address '$email' not valid."
                    }
                }
                #endregion

                #region Test Bcc
                foreach ($email in $Bcc) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "Bcc email address '$email' not valid."
                    }
                }
                #endregion

                #region Load MimeKit assembly
                if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                    try {
                        Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                        Add-Type -Path $MimeKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                    }
                }
                #endregion

                #region Load MailKit assembly
                if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                    try {
                        Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                        Add-Type -Path $MailKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                    }
                }
                #endregion
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
        }

        process {
            try {
                $message = New-Object -TypeName 'MimeKit.MimeMessage'

                #region Create body with attachments
                $bodyPart = New-Object MimeKit.TextPart('html')
                $bodyPart.Text = $Body

                $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
                $bodyMultiPart.Add($bodyPart)

                if ($Attachments) {
                    $params = @{
                        Attachments   = $Attachments
                        BodyMultiPart = $bodyMultiPart
                    }
                    $addAttachments = Add-Attachments @params

                    if ($addAttachments.AttachmentLimitExceededMessage) {
                        $bodyPart.Text += '<p><i>{0}</i></p>' -f
                        $addAttachments.AttachmentLimitExceededMessage
                    }
                }

                $message.Body = $bodyMultiPart
                #endregion

                $fromAddress = New-Object MimeKit.MailboxAddress(
                    $FromDisplayName, $From
                )
                $message.From.Add($fromAddress)

                foreach ($email in $To) {
                    $message.To.Add($email)
                }

                foreach ($email in $Bcc) {
                    $message.Bcc.Add($email)
                }

                $message.Subject = $Subject

                #region Set priority
                switch ($Priority) {
                    'Low' {
                        $message.Headers.Add('X-Priority', '5 (Lowest)')
                        break
                    }
                    'Normal' {
                        $message.Headers.Add('X-Priority', '3 (Normal)')
                        break
                    }
                    'High' {
                        $message.Headers.Add('X-Priority', '1 (Highest)')
                        break
                    }
                    default {
                        throw "Priority type '$_' not supported"
                    }
                }
                #endregion

                $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

                try {
                    $smtp.Connect(
                        $SmtpServerName, $SmtpPort,
                        [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
                    )
                }
                catch {
                    throw "Failed to connect to SMTP server '$SmtpServerName' on port '$SmtpPort' with connection type '$SmtpConnectionType': $_"
                }

                if ($Credential) {
                    try {
                        $smtp.Authenticate(
                            $Credential.UserName,
                            $Credential.GetNetworkCredential().Password
                        )
                    }
                    catch {
                        throw "Failed to authenticate with user name '$($Credential.UserName)' to SMTP server '$SmtpServerName': $_"
                    }
                }

                Write-Verbose "Send mail to '$To' with subject '$Subject'"

                $null = $smtp.Send($message)
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
            finally {
                if ($smtp) {
                    $smtp.Disconnect($true)
                    $smtp.Dispose()
                }
                if ($message) {
                    $message.Dispose()
                }
            }
        }
    }

    function Write-EventsToEventLogHC {
        <#
        .SYNOPSIS
            Write events to the event log.

        .DESCRIPTION
            The use of this function will allow standardization in the Windows
            Event Log by using the same EventID's and other properties across
            different scripts.

            Custom Windows EventID's based on the PowerShell standard streams:

            PowerShell Stream     EventIcon    EventID   EventDescription
            -----------------     ---------    -------   ----------------
            [i] Info              [i] Info     100       Script started
            [4] Verbose           [i] Info     4         Verbose message
            [1] Output/Success    [i] Info     1         Output on success
            [3] Warning           [w] Warning  3         Warning message
            [2] Error             [e] Error    2         Fatal error message
            [i] Info              [i] Info     199       Script ended successfully

        .PARAMETER Source
            Specifies the script name under which the events will be logged.

        .PARAMETER LogName
            Specifies the name of the event log to which the events will be
            written. If the log does not exist, it will be created.

        .PARAMETER Events
            Specifies the events to be written to the event log. This should be
            an array of PSCustomObject with properties: Message, EntryType, and
            EventID.

        .PARAMETER Events.xxx
            All properties that are not 'EntryType' or 'EventID' will be used to
            create a formatted message.

        .PARAMETER Events.EntryType
            The type of the event.

            The following values are supported:
            - Information
            - Warning
            - Error
            - SuccessAudit
            - FailureAudit

            The default value is Information.

        .PARAMETER Events.EventID
            The ID of the event. This should be a number.
            The default value is 4.

        .EXAMPLE
            $eventLogData = [System.Collections.Generic.List[PSObject]]::new()

            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script started'
                    EntryType = 'Information'
                    EventID   = '100'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message  = 'Failed to read the file'
                    FileName = 'C:\Temp\test.txt'
                    DateTime = Get-Date
                    EntryType = 'Error'
                    EventID   = '2'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message  = 'Created file'
                    FileName = 'C:\Report.xlsx'
                    FileSize = 123456
                    DateTime = Get-Date
                    EntryType = 'Information'
                    EventID   = '1'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script finished'
                    EntryType = 'Information'
                    EventID   = '199'
                }
            )

            $params = @{
                Source  = 'Test (Brecht)'
                LogName = 'HCScripts'
                Events  = $eventLogData
            }
            Write-EventsToEventLogHC @params
        #>

        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [String]$Source,
            [Parameter(Mandatory)]
            [String]$LogName,
            [PSCustomObject[]]$Events
        )

        try {
            if (
                -not(
                    ([System.Diagnostics.EventLog]::Exists($LogName)) -and
                    [System.Diagnostics.EventLog]::SourceExists($Source)
                )
            ) {
                Write-Verbose "Create event log '$LogName' and source '$Source'"
                New-EventLog -LogName $LogName -Source $Source -ErrorAction Stop
            }

            foreach ($eventItem in $Events) {
                $params = @{
                    LogName     = $LogName
                    Source      = $Source
                    EntryType   = $eventItem.EntryType
                    EventID     = $eventItem.EventID
                    Message     = ''
                    ErrorAction = 'Stop'
                }

                if (-not $params.EntryType) {
                    $params.EntryType = 'Information'
                }
                if (-not $params.EventID) {
                    $params.EventID = 4
                }

                foreach (
                    $property in
                    $eventItem.PSObject.Properties | Where-Object {
                        ($_.Name -ne 'EntryType') -and ($_.Name -ne 'EventID')
                    }
                ) {
                    $params.Message += "`n- $($property.Name) '$($property.Value)'"
                }

                Write-Verbose "Write event to log '$LogName' source '$Source' message '$($params.Message)'"

                Write-EventLog @params
            }
        }
        catch {
            throw "Failed to write to event log '$LogName' source '$Source': $_"
        }
    }

    try {
        $settings = $jsonFileContent.Settings

        $scriptName = $settings.ScriptName
        $saveInEventLog = $settings.SaveInEventLog
        $sendMail = $settings.SendMail
        $saveLogFiles = $settings.SaveLogFiles

        $allLogFilePaths = @()
        $baseLogName = $null
        $logFolderPath = $null

        #region Counter
        $counter = @{
            totalFilesCopied    = 0
            jobErrors           = ($Tasks.job.Error | Measure-Object).Count
            robocopyBadExitCode = 0
            robocopyJobError    = 0
            systemErrors        = $systemErrors.Count
            totalErrors         = 0
        }
        #endregion

        #region Create log folder
        try {
            $logFolder = Get-StringValueHC $saveLogFiles.Where.Folder

            $isLog = @{
                systemErrors = $saveLogFiles.What.systemErrors
                RobocopyLogs = $saveLogFiles.What.RobocopyLogs
            }

            if ($logFolder) {
                #region Get log folder
                try {
                    $logFolderPath = Get-LogFolderHC -Path $logFolder

                    Write-Verbose "Log folder '$logFolderPath'"

                    $baseLogName = Join-Path -Path $logFolderPath -ChildPath (
                        '{0} - {1} ({2})' -f
                        $scriptStartTime.ToString('yyyy_MM_dd'),
                        $ScriptName,
                        $jsonFileItem.BaseName
                    )
                }
                catch {
                    throw "Failed creating log folder '$LogFolder': $_"
                }
                #endregion

                #region Create log file
                if ($isLog.systemErrors -and $systemErrors) {
                    $params = @{
                        DataToExport   = $systemErrors
                        PartialPath    = "$baseLogName - System errors log"
                        FileExtensions = '.json'
                        Append         = $true
                    }
                    $allLogFilePaths += Out-LogFileHC @params
                }
                #endregion
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed creating log file in folder '$($saveLogFiles.Where.Folder)': $_"
                }
            )

            Write-Warning $systemErrors[-1].Message
        }
        #endregion

        $color = @{
            NoCopy   = 'White'     # Nothing copied
            CopyOk   = 'LightGrey' # Copy successful
            Mismatch = 'Orange'    # Clean-up needed (Mismatch)
            Fatal    = 'Red'       # Fatal error
        }

        $htmlTableRows = @()
        $i = 0

        Foreach (
            $job in
            $Tasks.Job.Results | Where-Object { $_ }
        ) {
            try {
                #region Verbose
                $M = "Job result: Name '$($job.Name)', ComputerName '$($job.ComputerName)', Source '$($job.Source)', Destination '$($job.Destination)', File '$($job.File)', Switches '$($job.Switches)', ExitCode '$($job.ExitCode)', Error '$($job.Error)'"

                Write-Verbose $M

                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = $M
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '2'
                    }
                )
                #endregion

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
                $logFile = $null

                if ($isLog.RobocopyLogs -and $logFolder) {
                    $i++

                    $logFile = "$baseLogName - {0} ($i) - Log.txt" -f
                    $(
                        Get-ValidFileNameHC $(
                            if ($job.Name) {
                                $job.Name
                            }
                            elseif ($job.Destination) {
                                $job.Destination
                            }
                            elseif ($job.InputFile) {
                                Split-Path $job.InputFile -Leaf
                            }
                        )
                    )

                    Write-Verbose "Create robocopy log file '$logFile'"

                    $params = @{
                        FilePath = $logFile
                        Encoding = 'utf8'
                    }
                    $job.RobocopyOutput | Out-File @params

                    $allLogFilePaths += $logFile
                }
                #endregion

                #region Create HTML table rows
                $robocopyLog = Convert-RobocopyLogToObjectHC $job.RobocopyOutput

                $robocopy = @{
                    ExitMessage   = Convert-RobocopyExitCodeToStringHC -ExitCode $job.ExitCode
                    ExecutionTime = if ($robocopyLog.Times.Total) {
                        $robocopyLog.Times.Total
                    }
                    else { 'NA' }
                    FilesCopied   = [INT]$robocopyLog.Files.Copied
                }

                $counter.totalFilesCopied += $robocopy.FilesCopied

                $htmlTableRows += @"
<tr bgcolor="$rowColor" style="background:$rowColor;">
    <td id="TxtLeft">{0}<br>{1}{2}{3}{4}</td>
    <td id="TxtLeft">$($robocopy.ExitMessage + ' (' + $job.ExitCode + ')')</td>
    <td id="TxtCentered">$($robocopy.ExecutionTime)</td>
    <td id="TxtCentered">$($robocopy.FilesCopied)</td>
    <td id="TxtCentered">{5}</td>
</tr>
"@ -f
                $(
                    if ($job.Name) {
                        $job.Name
                    }
                    else {
                        $path = if ($job.InputFile) {
                            $job.InputFile
                        }
                        elseif ($job.Source -match '^\\\\') {
                            $job.Source
                        }
                        else {
                            $job.Source -Replace '^.{2}', (
                                '\\{0}\{1}$' -f
                                $job.ComputerName, $job.Source[0]
                            )
                        }
                        '<a href="{0}">{0}</a>' -f $path
                    }
                ),
                $(
                    if (-not ($job.Name -or $job.InputFile)) {
                        $path = if ($job.Destination -match '^\\\\') {
                            $job.Destination
                        }
                        else {
                            $job.Destination -Replace '^.{2}', (
                                '\\{0}\{1}$' -f
                                $job.ComputerName, $job.Destination[0]
                            )
                        }
                        '<a href="{0}">{0}</a>' -f $path
                    }
                ),
                $(
                    "<br>Switches: $($job.Switches)"
                ),
                $(
                    if ($job.File) {
                        "<br>File: $($job.File)"
                    }
                ),
                $(
                    if ($job.Error) {
                        "<br><b>$($job.Error)</b>"
                    }
                ),
                $(
                    if ($logFile) {
                        '<a href="{0}">{1}</a>' -f $logFile, 'Log'
                    }
                    else {
                        'NA'
                    }

                )
                #endregion
            }
            catch {
                $systemErrors.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Failed creating robocopy log file or html table rows for '$M': $_"
                    }
                )

                Write-Warning $systemErrors[-1].Message
            }
        }

        #region Get script name
        if (-not $scriptName) {
            Write-Warning "No 'Settings.ScriptName' found in import file."
            $scriptName = 'Default script name'
        }
        #endregion

        #region Remove old log files
        if ($saveLogFiles.DeleteLogsAfterDays -gt 0 -and $logFolderPath) {
            $cutoffDate = (Get-Date).AddDays(-$saveLogFiles.DeleteLogsAfterDays)

            Write-Verbose "Removing log files older than $cutoffDate from '$logFolderPath'"

            Get-ChildItem -Path $logFolderPath -File |
            Where-Object { $_.LastWriteTime -lt $cutoffDate } |
            ForEach-Object {
                try {
                    $fileToRemove = $_
                    Write-Verbose "Deleting old log file '$_''"
                    Remove-Item -Path $_.FullName -Force
                }
                catch {
                    $systemErrors.Add(
                        [PSCustomObject]@{
                            DateTime = Get-Date
                            Message  = "Failed to remove file '$fileToRemove': $_"
                        }
                    )

                    Write-Warning $systemErrors[-1].Message

                    if ($baseLogName -and $isLog.systemErrors) {
                        $params = @{
                            DataToExport   = $systemErrors[-1]
                            PartialPath    = "$baseLogName - Errors"
                            FileExtensions = '.txt'
                        }
                        $allLogFilePaths += Out-LogFileHC @params -EA Ignore
                    }
                }
            }
        }
        #endregion

        #region Write events to event log
        try {
            $saveInEventLog.LogName = Get-StringValueHC $saveInEventLog.LogName

            if ($saveInEventLog.Save -and $saveInEventLog.LogName) {
                $systemErrors | ForEach-Object {
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            Message   = $_.Message
                            DateTime  = $_.DateTime
                            EntryType = 'Error'
                            EventID   = '2'
                        }
                    )
                }

                $eventLogData.Add(
                    [PSCustomObject]@{
                        Message   = 'Script ended'
                        DateTime  = Get-Date
                        EntryType = 'Information'
                        EventID   = '199'
                    }
                )

                $params = @{
                    Source  = $scriptName
                    LogName = $saveInEventLog.LogName
                    Events  = $eventLogData
                }
                Write-EventsToEventLogHC @params

            }
            elseif ($saveInEventLog.Save -and (-not $saveInEventLog.LogName)) {
                throw "Both 'Settings.SaveInEventLog.Save' and 'Settings.SaveInEventLog.LogName' are required to save events in the event log."
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed writing events to event log: $_"
                }
            )

            Write-Warning $systemErrors[-1].Message

            if ($baseLogName -and $isLog.systemErrors) {
                $params = @{
                    DataToExport   = $systemErrors[-1]
                    PartialPath    = "$baseLogName - Errors"
                    FileExtensions = '.txt'
                }
                $allLogFilePaths += Out-LogFileHC @params -EA Ignore
            }
        }
        #endregion

        #region Create error log file
        if ($isLog.systemErrors -and $systemErrors -and $baseLogName) {
            $params = @{
                DataToExport   = $systemErrors
                PartialPath    = "$baseLogName - System errors log"
                FileExtensions = '.txt'
                Append         = $true
            }
            $allLogFilePaths += Out-LogFileHC @params
        }

        if ($isLog.systemErrors -and $counter.jobErrors -and $baseLogName) {
            $params = @{
                DataToExport   = $Tasks | Where-Object { $_.Job.Error }
                PartialPath    = "$baseLogName - System errors log"
                FileExtensions = '.txt'
                Append         = $true
            }
            $allLogFilePaths += Out-LogFileHC @params
        }
        #endregion

        #region Send email
        try {
            $isSendMail = switch ($sendMail.When) {
                'Never' { $false }
                'Always' { $true }
                'OnError' { $counter.totalErrors -gt 0 }
                'OnErrorOrAction' {
                    ($counter.totalErrors -gt 0) -or
                    ($counter.totalFilesCopied -gt 0)
                }
                default {
                    throw "SendMail.When '$($sendMail.When)' not supported. Supported values are 'Never', 'Always', 'OnError' or 'OnErrorOrAction'."
                }
            }

            if ($isSendMail) {
                #region Test mandatory fields
                @{
                    'From'                 = $sendMail.From
                    'Smtp.ServerName'      = $sendMail.Smtp.ServerName
                    'Smtp.Port'            = $sendMail.Smtp.Port
                    'AssemblyPath.MailKit' = $sendMail.AssemblyPath.MailKit
                    'AssemblyPath.MimeKit' = $sendMail.AssemblyPath.MimeKit
                }.GetEnumerator() |
                Where-Object { -not $_.Value } | ForEach-Object {
                    throw "Input file property 'Settings.SendMail.$($_.Key)' cannot be blank"
                }
                #endregion

                #region Create HTML table
                $htmlTable = $null

                if ($htmlTableRows) {
                    $htmlTable = @"
            <table id="TxtLeft">
                <tr>
                    <th id="TxtLeft">Robocopy</th>
                    <th id="TxtLeft">Status</th>
                    <th id="TxtCentered" class="Centered">Duration</th>
                    <th id="TxtCentered" class="Centered">Items</th>
                    <th id="TxtCentered" class="Centered">Logs</th>
                </tr>
                $htmlTableRows
            </table>
            <br>
            <table id="LegendTable">
                 <tr>
                    <td bgcolor="$($color.NoCopy)" style="background:$($color.NoCopy);" id="LegendRow">Nothing copied</td>
                    <td bgcolor="$($color.CopyOk)" style="background:$($color.CopyOk);" id="LegendRow">Copy successful</td>
                    <td bgcolor="$($color.Mismatch)" style="background:$($color.Mismatch);" id="LegendRow">Clean-up needed</td>
                    <td bgcolor="$($color.Fatal)" style="background:$($color.Fatal);" id="LegendRow">Fatal error</td>
                </tr>
            </table>
"@
                }
                #endregion

                $mailParams = @{
                    From                = Get-StringValueHC $sendMail.From
                    SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
                    SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
                    MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
                    MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
                    Subject             = '{0} job{1}, {2} item{3}' -f
                    $Tasks.Count,
                    $(if ($Tasks.Count -ne 1) { 's' }),
                    $counter.totalFilesCopied,
                    $(if ($counter.totalFilesCopied -ne 1) { 's' })
                }

                #region Set mail subject and priority
                if (
                    $counter.totalErrors = $counter.systemErrors + $counter.jobErrors +
                    $counter.robocopyBadExitCode + $counter.robocopyJobError
                ) {
                    $mailParams.Subject += ', {0} error{1}' -f
                    $counter.totalErrors, $(if ($counter.totalErrors -ne 1) { 's' })
                    $mailParams.Priority = 'High'
                }
                #endregion

                $mailParams.Body = @"
<!DOCTYPE html>
<html>
<head>
<style type="text/css">
    body {
        font-family:verdana;
        font-size:14px;
        background-color:white;
    }
    h1 {
        margin-bottom: 0;
    }
    h2 {
        margin-bottom: 0;
    }
    h3 {
        margin-bottom: 0;
    }
    p.italic {
        font-style: italic;
        font-size: 12px;
    }
    table {
        border-collapse:collapse;
        border:0px none;
        padding:3px;
        text-align:left;
    }
    td, th {
        border-collapse:collapse;
        border:1px none;
        padding:3px;
        text-align:left;
    }
    #aboutTable th {
        color: rgb(143, 140, 140);
        font-weight: normal;
    }
    #aboutTable td {
        color: rgb(143, 140, 140);
        font-weight: normal;
    }
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
    base {
        target="_blank"
    }
</style>
</head>
<body>
<table>
    <h1>$scriptName</h1>
    <hr size="2" color="#06cc7a">

    $($sendMail.Body)

    <table>
        <tr>
            <th>Tasks</th>
            <td>$($Tasks.Count)</td>
        </tr>
        <tr>
            <th>Items</th>
            <td>$($counter.totalFilesCopied)</td>
        </tr>
        $(
            $counter.systemErrors ?
            "<tr style=`"background-color: #ffe5ec;`">
                <th>System errors</th>
                <td>$($counter.systemErrors)</td>
            </tr>" : ''
        )
        $(
            $counter.jobErrors ?
            "<tr style=`"background-color: #ffe5ec;`">
                <th>Job errors</th>
                <td>$($counter.jobErrors)</td>
            </tr>" : ''
        )
        $(
            $counter.robocopyBadExitCode ?
            "<tr style=`"background-color: #ffe5ec;`">
                <th>Tasks with errors in robocopy log files</th>
                <td>$($counter.robocopyBadExitCode)</td>
            </tr>" : ''
        )
        $(
            $counter.robocopyJobError ?
            "<tr style=`"background-color: #ffe5ec;`">
                <th>Errors while executing robocopy</th>
                <td>$($counter.robocopyJobError)</td>
            </tr>" : ''
        )
    </table>

    $htmlTable

    $(
        if ($allLogFilePaths) {
            '<p><i>* Check the attachment(s) for details</i></p>'
        }
    )

    <hr size="2" color="#06cc7a">
    <table id="aboutTable">
        $(
            if ($scriptStartTime) {
                '<tr>
                    <th>Start time</th>
                    <td>{0:00}/{1:00}/{2:00} {3:00}:{4:00} ({5})</td>
                </tr>' -f
                $scriptStartTime.Day,
                $scriptStartTime.Month,
                $scriptStartTime.Year,
                $scriptStartTime.Hour,
                $scriptStartTime.Minute,
                $scriptStartTime.DayOfWeek
            }
        )
        $(
            if ($scriptStartTime) {
                $runTime = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
                '<tr>
                    <th>Duration</th>
                    <td>{0:00}:{1:00}:{2:00}</td>
                </tr>' -f
                $runTime.Hours, $runTime.Minutes, $runTime.Seconds
            }
        )
        $(
            if ($logFolderPath) {
                '<tr>
                    <th>Log files</th>
                    <td><a href="{0}">Open log folder</a></td>
                </tr>' -f $logFolderPath
            }
        )
        <tr>
            <th>Host</th>
            <td>$($host.Name)</td>
        </tr>
        <tr>
            <th>PowerShell</th>
            <td>$($PSVersionTable.PSVersion.ToString())</td>
        </tr>
        <tr>
            <th>Computer</th>
            <td>$env:COMPUTERNAME</td>
        </tr>
        <tr>
            <th>Account</th>
            <td>$env:USERDNSDOMAIN\$env:USERNAME</td>
        </tr>
    </table>
</table>
</body>
</html>
"@

                if ($sendMail.FromDisplayName) {
                    $mailParams.FromDisplayName = Get-StringValueHC $sendMail.FromDisplayName
                }

                if ($sendMail.Subject) {
                    $mailParams.Subject = '{0}, {1}' -f
                    $mailParams.Subject, $sendMail.Subject
                }

                if ($sendMail.To) {
                    $mailParams.To = $sendMail.To
                }

                if ($sendMail.Bcc) {
                    $mailParams.Bcc = $sendMail.Bcc
                }

                if ($allLogFilePaths) {
                    $mailParams.Attachments = $allLogFilePaths |
                    Sort-Object -Unique
                }

                if ($sendMail.Smtp.ConnectionType) {
                    $mailParams.SmtpConnectionType = Get-StringValueHC $sendMail.Smtp.ConnectionType
                }

                #region Create SMTP credential
                $smtpUserName = Get-StringValueHC $sendMail.Smtp.UserName
                $smtpPassword = Get-StringValueHC $sendMail.Smtp.Password

                if ( $smtpUserName -and $smtpPassword) {
                    try {
                        $securePassword = ConvertTo-SecureString -String $smtpPassword -AsPlainText -Force

                        $credential = New-Object System.Management.Automation.PSCredential($smtpUserName, $securePassword)

                        $mailParams.Credential = $credential
                    }
                    catch {
                        throw "Failed to create credential: $_"
                    }
                }
                elseif ($smtpUserName -or $smtpPassword) {
                    throw "Both 'Settings.SendMail.Smtp.Username' and 'Settings.SendMail.Smtp.Password' are required when authentication is needed."
                }
                #endregion

                Send-MailKitMessageHC @mailParams
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed sending email: $_"
                }
            )

            Write-Warning $systemErrors[-1].Message

            if ($baseLogName -and $isLog.systemErrors) {
                $params = @{
                    DataToExport   = $systemErrors[-1]
                    PartialPath    = "$baseLogName - System errors"
                    FileExtensions = '.txt'
                }
                $null = Out-LogFileHC @params -EA Ignore
            }
        }
        #endregion
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = $_
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
    finally {
        if ($systemErrors) {
            $M = 'Found {0} system error{1}' -f
            $systemErrors.Count,
            $(if ($systemErrors.Count -ne 1) { 's' })
            Write-Warning $M

            $systemErrors | ForEach-Object {
                Write-Warning $_.Message
            }

            Write-Warning 'Exit script with error code 1'
            exit 1
        }
        else {
            Write-Verbose 'Script finished successfully'
        }
    }
}