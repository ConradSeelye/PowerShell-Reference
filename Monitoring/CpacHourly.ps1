<#
.SYNOPSIS
==============================================================================
                         ________  _______  _   __
                        / ____/  |/  / __ \/ | / /
                       / / __/ /|_/ / / / /  |/ /
                      / /_/ / /  / / /_/ / /|  /
                      \____/_/  /_/\____/_/ |_/

==============================================================================
CpacHourly.ps1					                       
		 (Supports CPAC performance data collection)

.DESCRIPTION
      Cycles the CPAC perfmon logs, archiving logs over a certain number of days
 age, deletes archived logs over a certain # of days age, runs a system
 configuration inventory, and invokes an upload transfer script to send performance
 data centrally.


.NOTES
OWNER: Gillette, Steven A - Windows Server Technology
---[ Revision ]---------------------------------------------------------------
   S.A. Gillette       12-10-25     Enhancements for GMON v6
                                       -Add Option for random startup delay secs (script param)
                                       -Add option for pre-typeperf command delay (CpacHourly.cfgx option)                                         
   S.A. Gillette       11-29-22     Enhancements for GMON v5
                                       -Add CpacHourlyCfg.json into .zip files for central visibility of upload url/settings
                                       -Immediately abort if .\LibGmonAgent.ps1 cannot be loaded (Contains required cleanup functions)                                      
   S.A. Gillette       1-3-19       Bug #CPAC-418: GMON v3 not creating CSVs in the daily archive zips
   S.A. Gillette       5-23-18      -3.0 development of a non-combined BLG approach to workaround relog bugs
                                    Bug #CPAC-336: GMON sends improperly formatted dates to relog in non-US locales
                                    Bug #CPAC-220: Corrupt Process Level Metrics for GMON
   S.A. Gillette       2-2-18       -BUGID #11: UCased $ComputerName for better filename consistency
   S.A. Gillette       1-26-18   v1.0.2 - Several enhancements/bug fixes:
                                    -BUGID #1: Do not generate Hourly archive zip if RELOG fails
                                               to produce a CSV log (as it produces a zip with no perf data)
                                    -BUGID #2: Not skipping log processing after reboot
                                    -BUGID #3: Size calculations generates runtime errors when 
                                               no files are present to measure
                                    -BUGID #4: Script default upload delay was 240s instead of 180s
                                    -BUGID #9: Modularity: Move functions to library, use calls to 
                                               Remove-OldFiles istead of repeated cleanup code sections
   S.A. Gillette       5-24-17    v1.0.1 - Bug fix for hour 23 not getting logged.
   S.A. Gillette       10-17-16   -New script for new CPAC perfmon agent

==============================================================================
#>

[CmdletBinding()]
Param (
    # A randomized startup delay window before the script performs any activity
    # (Can help avoid impacts from multiple VMs starting activity simulataneously)
    # i.e. a value of 7 will result in randomly delaying anywhere from 0-7 seconds)
    [Parameter(Mandatory=$False, Position=0)]
    [ValidateRange(0,3300)]
    [Uint32] $StartupDelaySecs = 0
)

#region functions
# ------------------------------------------------------------------------------
# FUNCTIONS
# ------------------------------------------------------------------------------


#endregion functions


#region scriptinitialization
# ------------------------------------------------------------------------------
$ScriptStarted = Get-Date

# Begin startup delay if configured
if ($StartupDelaySecs -gt 0) {
    $StartupDelayMs = Get-Random -Minimum 0 -Maximum ($StartupDelaySecs * 1000)
    "Delaying script startup for 0-$StartupDelaySecs secs ($($StartupDelayMs.ToString('n0')) ms actual)..."
    Start-Sleep -Milliseconds $StartupDelayMs
} Else {
    "Continuing without startup delay."
    $StartupDelayMs = 0 # this gets used in a msg later
}

# Load functions from libary.
"Loading LibGmonAgent.ps1 function library..."
. .\LibGmonAgent.ps1
if (-not $?) {
    throw "Error loading required LibGmonAgent.ps1 function library!  Aborting!"
}
$HourOfDay = $ScriptStarted.Hour
$DateStamp = $ScriptStarted.ToString("yyyyMMdd")
$ComputerName = (Get-Item Env:\COMPUTERNAME).Value.ToUpper()
$ScriptPath = (Get-Item $MyInvocation.InvocationName).FullName
$ScriptDir = (Get-Item $ScriptPath).Directory.FullName
$LogDir = (Get-Item $ScriptPath).Directory.Parent.FullName + "\logs"
$InventoryDir = (Get-Item $ScriptPath).Directory.Parent.FullName + "\machineinfo"
$ConfigDir = (Get-Item $ScriptPath).Directory.Parent.FullName + "\etc"
$FilteredHourlyArchiveDir = "$LogDir\archive\filteredhourly"
$FilteredDailyArchiveDir = "$LogDir\archive\filtereddaily"
$GranularArchiveDir = "$LogDir\archive\granular"
$FreeDiskGiB = [math]::Round(((Get-PSDrive 'C').Free / 1GB), 2)

$TimedOps = @()
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Script Started with a delay of 0-$StartupDelaySecs secs ($($StartupDelayMs.ToString('n0')) ms actual)"
# start transcript
$TranscriptLog = "$LogDir\" + (Get-Item $MyInvocation.InvocationName).BaseName + "_$DateStamp.log"
Start-Transcript "$TranscriptLog" -Append
If ($StartupDelayMs) {
    Write-Warning "Script startup was randomly delayed 0-$StartupDelaySecs secs ($($StartupDelayMs.ToString('n0')) ms actual)"
}
$ERROR.Clear()
Clear-Host
# -----------------------------

#$DayName = $ScriptStarted.DayOfWeek
#$DayOfMonth = $ScriptStarted.Month
$Today = Get-Date "0:00"
$Yesterday = $Today.AddDays(-1)
$ThisHour = $Today.AddHours($HourOfDay)
$LastHour = $Today.AddHours($HourOfDay - 1)
$SystemBootTime = (GWMI WIn32_OperatingSystem | Select-Object @{Name="LastBootUpTime";Expression={$_.ConverttoDateTime($_.lastbootuptime)}}).LastBootUpTime
$UpTimeMins = [math]::Round(($ScriptStarted - $SystemBootTime).TotalMinutes, 2)
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Initialization Completed"

#Check for CollectionDisabled flag file.
# If this file exists in ..\etc then script aborts immediately.
$DisabledFlagFile = "$ConfigDir\CpacHourly.DISABLED"
if (Test-Path $DisabledFlagFile) {
    " # WARNING: Collection has been disabled due to file $DisabledFlagFile!  Terminating!"
    Stop-Transcript
    Exit 1818 # ~ Operation Cancelled
}

#endregion scriptinitialization

#region loadparameters
#CPAC Daily Default Params
$DefaultOptions = New-Object PSObject
$DefaultOptions | Add-Member -Type NoteProperty -Name BlgSampleIntervalSecs -value 60
$DefaultOptions | Add-Member -Type NoteProperty -Name TypeperfDelaySecs -value 0
$DefaultOptions | Add-Member -Type NoteProperty -Name CpacDesampleRatio -value 5
$DefaultOptions | Add-Member -Type NoteProperty -Name LogRetainDays -value 2  #Affects all files in ..\logs
$DefaultOptions | Add-Member -Type NoteProperty -Name MinFreeDiskGiB -value 5 #Collection will abort if less than this free disk space left.
$DefaultOptions | Add-Member -Type NoteProperty -Name ArchiveFilteredHourlyRetainDays -value 2
$DefaultOptions | Add-Member -Type NoteProperty -Name ArchiveFilteredDailyRetainDays -value 30
$DefaultOptions | Add-Member -Type NoteProperty -Name ArchiveGranularDailyRetainDays -value 7
$DefaultOptions | Add-Member -Type NoteProperty -Name DailyCycleHour -value 0 #What hour of day the daily zips are produced
$DefaultOptions | Add-Member -Type NoteProperty -Name InventoryEnabled -value $True  # Set to false to disable inventory
$DefaultOptions | Add-Member -Type NoteProperty -Name InventoryHour -value 23 #Ideally, set 1 hour before DailyCycleHour
$DefaultOptions | Add-Member -Type NoteProperty -Name UploadEnabled -value $True # If false, no uploads triggerd
$DefaultOptions | Add-Member -Type NoteProperty -Name UploadToCpacHourly -value $True  # If false, upload daily
$DefaultOptions | Add-Member -Type NoteProperty -Name UploadDelayEnabled -value $True
$DefaultOptions | Add-Member -Type NoteProperty -Name UploadDelaySecs -value (60 * 3)  # 3 mins (randomized)
$DefaultOptions | Add-Member -Type NoteProperty -Name UploadServiceUri -value 'https://gmon.web.boeing.com/Upload/Upload.asmx?WSDL'

$ParamFile = "$ConfigDir\CpacHourlyCfg.json"
"[$(Get-Date -f T)] Loading configuration options from '$ParamFile'"
If (Test-Path $ParamFile) {
    $FileOptions = Get-Content $ParamFile -Encoding UTF8 | Out-String | ConvertFrom-Json
} else {
    "NOTE: Unable to find $ParamFile; Creating a new one using script defaults."
    $FileOptions = $DefaultOptions
    $FileOptions | ConvertTo-Json | Out-file $ParamFile -Encoding utf8
}
# Resolving runtime options

"[$(Get-Date -f T)] Resolving runtime parameter defaults vs those in $ParamFile..."
$RuntimeOptions = $DefaultOptions # Until a value is overridden
$Comparison = Compare-Object -ReferenceObject ($DefaultOptions | Get-Member -MemberType NoteProperty) `
    -DifferenceObject ($FileOptions | Get-Member -MemberType NoteProperty) `
    -IncludeEqual -Property Name
$Comparison | ForEach-Object {
    $PropName = $_.Name
    If ($_.SideIndicator -eq "=>") {
        #Only defined in Option File
        Write-Warning "The property '$PropName' defined in $ParamFile is unrecognized and will be ignored!"
    } ElseIf ($_.SideIndicator -eq "<=") {
        # Not defined in config file, use default value
        #$RuntimeOptions | Add-Member -Type NoteProperty -Name "$PropName" -value $DefaultOptions."$PropName"
    } Else {
        #Both files have same property, is the value of it the same?
        # (Use File option if values are different from default)
        If ($DefaultOptions."$PropName" -ne $FileOptions."$PropName") {
            "   NOTE: $PropName will be overridden to $($FileOptions."$PropName") from default of $($DefaultOptions."$PropName")." 
            $RuntimeOptions."$PropName" = $FileOptions."$PropName"
        }
    }
}
"Runtime Options have been set to:"
$RuntimeOptions | Format-List
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Determined runtime Parameters"
#endregion loadparameters

#Free Disk space check
If ($FreeDiskGiB -lt $RuntimeOptions.MinFreeDiskGiB) {
    Write-Error "Aborting collection because there is only $FreeDiskGiB GiB of disk space left on C:!"
    $StartPerfLog = $False
} else {
    "[$(Get-Date -f T)] SUCCESS: There is $FreeDiskGiB GiB of disk space on C:-- proceeding with collection."
    $StartPerfLog = $True
}

#region TypeperfDelay
# If the option to delay typeperf operations is configured, determine/perform a random wait.
if ($RuntimeOptions.TypeperfDelaySecs -gt 0) {
    $TypePerfDelayMs = Get-Random -Minimum 0 -Maximum ($RuntimeOptions.TypeperfDelaySecs * 1000)
    "Delaying typeperf.exe commands for 0-$($RuntimeOptions.TypeperfDelaySecs) secs ($($TypePerfDelayMs.ToString('n0')) ms actual)..."
    Start-Sleep -Milliseconds $TypePerfDelayMs
    $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Delayed Typeperf operations 0-$($RuntimeOptions.TypeperfDelaySecs) secs ($($TypePerfDelayMs.ToString('n0')) ms actual)"
}
#endregion TypeperfDelay

#region restartlogging
"[$(Get-Date -f T)] Stopping running CPAC typeperf.exe processes from previous collection cycle(s)..." 
$TypePerfProcs = @(Get-WmiObject Win32_Process | Where-Object {$_.Name -eq 'typeperf.exe' -and $_.CommandLine -match "CPAC"})
foreach ($Process in $TypePerfProcs) {
    "Stopping $($Process.Name) PID# $($Process.ProcessID)"
    Stop-Process -Id $Process.ProcessID -Force
}
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Stopped Typeperf"

""

# Figure out how many counter samples to log until the end of the hour
$TopOfHour = Get-Date (Get-Date).AddHours(1).ToString("MMM dd, yyyy HH:00")
$SecsToTopOfHour = ($TopOfHour - (Get-Date)).TotalSeconds
$SampleIntervalSecs = $RuntimeOptions.BlgSampleIntervalSecs
$SampleCount = [math]::Floor($SecsToTopOfHour / $SampleIntervalSecs)
$CounterFile = "$ConfigDir\counters_granular.txt"

$RunningBlgLog = "$LogDir\$ComputerName`__$(Get-Date -f "yyyyMMdd-HHmmss").blg"
$ArgList = "-cf `"$CounterFile`" -sc $SampleCount -si $SampleIntervalSecs -o `"$RunningBlgLog`" -f blg -y"
$TypeperfStdOut = "$LogDir\typeperf.stdout.txt"
$TypeperfStdErr = "$LogDir\typeperf.stderr.txt"
If ($StartPerfLog) {
    "[$(Get-Date -f T)] Starting new typeperf process for next collection cycle..."
    

    "Start-Process typeperf.exe $ArgList  -RedirectStandardError $TypeperfStdErr -RedirectStandardOutput $TypeperfStdOut"
    Start-Process typeperf.exe $ArgList -RedirectStandardError $TypeperfStdErr -RedirectStandardOutput $TypeperfStdOut
    if ($?) {
        "[$(Get-Date -f T)] Process started.  stdOut and stdErr are redirected to files shown above"
        "Note: There were ($SampleCount) $SampleIntervalSecs-sec intervals to log until $($TopOfHour.ToString('G'))"
    } Else {
        Write-Error "Failed to start process. Try checking stdOut/stdErr files above for more details"
    }
    $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Restarted Typeperf"
} Else {
    "[$(Get-Date -f T)] WARNING: No new log was started (due to insufficient disk space).  Cleanup and inventory will still be performed, however."
}
#endregion restartlogging


# Don't process logs if the system was just rebooted
# Perform an immediate inventory if system was just booted, or has no config.json.
#    Otherwise Perform a delayed inventory once per day during the inventory hour.
"[$(Get-Date -f T)] INFO: System was last booted: $SystemBootTime (Uptime: $UpTimeMins mins)"
If ($UpTimeMins -lt 5) {
    # System has just booted < 5 mins ago.
    Write-Warning "Skipping BLG processing because system was booted less than 5 mins ago."
    $ProcessLogs = $False
    "-----------------"
    If ($RuntimeOptions.InventoryEnabled) {
        Write-Warning "System was booted less than 5 mins ago so it will be re-inventoried immediately."
        $PerformDelayedInventory = $False  #
        "Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\ConfigJson.ps1"
        Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\ConfigJson.ps1
        "Script exited with code: $LASTEXITCODE"
        $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Performed immediate inventory (Due to recent boot)"
    } Else {
        "WARNING: Skipping inventory because InventoryEnabled=False"
        $PerformDelayedInventory = $False # Don't do a delayed one either.
    }
} Else {
    # System HAS been up for 5+ mins
    $InventoryHour = $RuntimeOptions.InventoryHour
    If (($HourOfDay -eq $InventoryHour) -or ($InventoryHour -eq "*")) {
        $PerformDelayedInventory = $True
        "[$(Get-Date -f T)] Current Hour ($HourOfDay) IS an inventory hour ($InventoryHour);  A system inventory will be PERFORMED."
    } ElseIf (-Not (Test-Path "$InventoryDir\Config.json")) {
        "[$(Get-Date -f T)] It is not the inventory hour, however, '$InventoryDir\Config.json' does not exist, so inventory will be run immediately"
        $PerformDelayedInventory = $False
        Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\ConfigJson.ps1
        $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Performed immediate inventory (None exists currently)"
    } Else {
        $PerformDelayedInventory = $False
        "[$(Get-Date -f T)] The Current Hour ($HourOfDay) is NOT an inventory hour ($InventoryHour);  A system inventory will be SKIPPED."
    }
    $ProcessLogs = $True
}

#region removeoldlogs
"# ---------------------"
"# Delete old log/blg files"
$LogRetainDays = $RuntimeOptions.LogRetainDays
If (($Null -eq $LogRetainDays) -or ($LogRetainDays -le 0)) { $LogRetainDays = 2} # Safety check to not delete ALL logs
Remove-OldFiles -Path $LogDir\*.* -Exclude $RunningBlgLog -MaxAge $LogRetainDays
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Removed Old BLG Files"

"# ---------------------"
"# Delete Hourly ZIP archives"
$MaxAge = $RuntimeOptions.ArchiveFilteredHourlyRetainDays
If (($Null -eq $MaxAge) -or ($MaxAge -le 0)) {$MaxAge = 2} # Safety check to not delete ALL logs
Remove-OldFiles -Path $FilteredHourlyArchiveDir\*.* -MaxAge $MaxAge
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Removed Hourly Filtered ZIP Files"

"# ---------------------"
"# Delete old granular archives"
$MaxAge = $RuntimeOptions.ArchiveGranularDailyRetainDays
If (($Null -eq $MaxAge) -or ($MaxAge -le 0)) {$MaxAge = 7} # Safety check to not delete ALL logs
if (-not (Test-path $GranularArchiveDir)) {mkdir$GranularArchiveDir}
Remove-OldFiles -Path "$GranularArchiveDir\*.*" -MaxAge $MaxAge
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Removed Hourly Granular ZIP Files"

"# Delete old cpacdaily ZIP archives"
$MaxAge = $RuntimeOptions.ArchiveFilteredDailyRetainDays
If (($Null -eq $MaxAge) -or ($MaxAge -le 0)) {$MaxAge = 30} # Safety check to not delete ALL logs
if (-not (Test-path $FilteredDailyArchiveDir)) {mkdir$FilteredDailyArchiveDir}
Remove-OldFiles -Path "$FilteredDailyArchiveDir\*.*" -MaxAge $MaxAge
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Removed Daily Filtered ZIP Files"
#endregion removeoldlogs


#region hourlyconvert
if (-not $ProcessLogs) {
    "# SKIPPING log processing because `$ProcessLogs is set to: $ProcessLogs"
} Else {
     #Proceed with log processing

    "#------------------------"
    "# Convert previous hour BLG files to CSV and create HOURLY FILTERED zip...."

    $LogBeginTS = "$($LastHour.ToString("d")) $($LastHour.ToString("HH:mm:ss"))"
    $LogEndTS = "$($ThisHour.AddSeconds(-1).ToString("d")) $($ThisHour.AddSeconds(-1).ToString("HH:mm:ss"))"
    $BlgsToConvert = @(Get-ChildItem "$LogDir\$ComputerName*.blg" -Exclude $RunningBlgLog | Where-Object {$_.LastWriteTime -ge $LastHour -and $_.CreationTime -lt $ThisHour.AddSeconds(-30)} | Sort-Object LastWriteTime)
    #Relog params
    $DesampleRatio = $RuntimeOptions.CpacDesampleRatio
    $CounterFile = $CounterFile = "$ConfigDir\counters_filtered.txt"
    If ($DesampleRatio -notmatch "^(\d+)$") {$DesampleRatio = 1} # 1 = no desampling
    # make dirs if needed and remove previous temp data
    If (-not (Test-Path "$LogDir\tmp")) {mkdir"$LogDir\tmp"}
    Remove-Item "$LogDir\tmp\*" -Recurse -Force -ErrorAction Ignore
    If (-not (Test-Path "$FilteredHourlyArchiveDir")) {mkdir"$FilteredHourlyArchiveDir"}
    
    "[$(Get-Date -f T)] Converting data from the following files into HOURLY FILTERED CSV:"
    $BlgsToConvert | ForEach-Object {$_.Name}
    $i = 0
    foreach ($BlgFile in $BlgsToConvert)
    {
        $i++
        $CsvFile = "$LogDir\tmp\$($BlgFile.BaseName).csv"
        "relog.exe $($BlgFile.FullName) -cf $CounterFile -f csv -o $CsvFile -b $LogBeginTS -e $LogEndTS -y -t $DesampleRatio"
        relog.exe $BlgFile.FullName -cf $CounterFile -f csv -o $CsvFile -b $LogBeginTS -e $LogEndTS -y -t $DesampleRatio
    }
    "[$(Get-Date -f T)] Finished converting $i BLG files into CSV"
    $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Converted ($i) HOURLY BLG files to CSV"
    #endregion hourlyconvert

    #region hourlyzip
    "#------------------------"
    "# Hourly filtered zip creation "
    if (Test-Path "$InventoryDir\Config.json") {
        Copy-Item "$InventoryDir\Config.json" "$LogDir\tmp\$ComputerName`_config.json" # put a copy of config.json in temp folder to be zipped
    }
    if (Test-Path $ParamFile) {
        Copy-Item $ParamFile "$LogDir\tmp\$ComputerName`_CpacHourlyCfg.json" # put a copy of cpachourlycfg.json in temp folder to be zipped
    }
    $HourlyZipFile = "$FilteredHourlyArchiveDir\$ComputerName`_$($LastHour.ToString("yyyyMMdd-HH"))" + ".zip"

    # Zip it!
    Zip-Folder -FolderPath "$LogDir\tmp" -ZipFilePath $HourlyZipFile -Overwrite | Format-Table -AutoSize
    $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Zippped Hourly Filtered data"
    #endregion hourlyzip



    #region dailyprocessing
    "#---------------------------------"
    "# Once-Daily processing tasks"
    If ($HourOfDay -ne $RuntimeOptions.DailyCycleHour) {
        "[$(Get-Date -f T)] Current Hour ($HourOfDay) is NOT the daily cycle hour ($($RuntimeOptions.DailyCycleHour)); SKIPPING daily cycle operations..."
    } Else {
        "[$(Get-Date -f T)] Current Hour ($HourOfDay) is the daily cycle hour ($($RuntimeOptions.DailyCycleHour)); performing daily cycle operations..."
        $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Starting Daily processing"
        #region dailygranularconvert
        "# Daily GRANULAR BLG Convert"
        $LogBeginTS = "$($Yesterday.ToString("d")) $($Yesterday.ToString("HH:mm:ss"))"
        $LogEndTS = "$($Today.AddSeconds(-1).ToString("d")) $($Today.AddSeconds(-1).ToString("HH:mm:ss"))"
        $CounterFile = $CounterFile = "$ConfigDir\counters_granular.txt"
        $BlgsToConvert = @(Get-ChildItem "$LogDir\$ComputerName*.blg" -Exclude $RunningBlgLog | Where-Object {$_.LastWriteTime -ge $Yesterday -and $_.CreationTime -lt $Today.AddSeconds(-30)} | Sort-Object LastWriteTime)
        Remove-Item "$LogDir\tmp\*" -Recurse -Force -ErrorAction Ignore

        "[$(Get-Date -f T)] Converting data from the following BLG files into GRANULAR CSV:"
        $BlgsToConvert | ForEach-Object {$_.Name}
        $i = 0
        foreach ($BlgFile in $BlgsToConvert)
        {
            $i++
            $CsvFile = "$LogDir\tmp\$($BlgFile.BaseName).csv"
            "relog.exe $($BlgFile.FullName) -cf $CounterFile -f csv -o $CsvFile -b $LogBeginTS -e $LogEndTS -y"
            relog.exe $BlgFile.FullName -cf $CounterFile -f csv -o $CsvFile -b $LogBeginTS -e $LogEndTS -y
            
        }
        "[$(Get-Date -f T)] Finished Converting $i BLG files into CSV"
        $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Finished Converting Daily BLG Files to CSV ($i)"
        #endregion dailygranularconvert

        #region dailygranularzip
        "#----------------------------------"
        "# Creating daily granular archive zip"

        Copy-Item "$InventoryDir\Config.json" "$LogDir\tmp\$ComputerName`_config.json" # put a copy of config.json in temp folder to be zipped
        if (Test-Path $ParamFile) {
            Copy-Item $ParamFile "$LogDir\tmp\$ComputerName`_CpacHourlyCfg.json" # put a copy of cpachourlycfg.json in temp folder to be zipped
        }

        #Zip it
        $ZipFileName = "$GranularArchiveDir\$ComputerName`_$($Yesterday.ToString("yyyyMMdd"))" + ".zip"
        "[$(Get-Date -f T)] Compressing log and configuration data to $ZipFileName"
        Zip-Folder -FolderPath "$LogDir\tmp" -ZipFilePath $ZipFileName -Overwrite | Format-Table -AutoSize
        $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Created daily granular ZIP file"
        #endregion dailygranularzip


        #region dailycpaczip
        "#----------------------------------"
        "# Create daily filtered zip"
        Remove-Item "$LogDir\tmp\*" -Recurse -Force -ErrorAction Ignore
        Copy-Item "$InventoryDir\Config.json" "$LogDir\tmp\$ComputerName`_config.json" # put a copy in temp folder to be zipped
        if (Test-Path $ParamFile) {
            Copy-Item $ParamFile "$LogDir\tmp\$ComputerName`_CpacHourlyCfg.json" # put a copy of cpachourlycfg.json in temp folder to be zipped
        }

        $DailyZipFile = "$FilteredDailyArchiveDir\$ComputerName`_$($Yesterday.ToString("yyyyMMdd"))" + ".zip"
        $DesampleRatio = $RuntimeOptions.CpacDesampleRatio
        If ($DesampleRatio -notmatch "^(\d+)$") {$DesampleRatio = 1}
        $CounterFile = $CounterFile = "$ConfigDir\counters_filtered.txt"

        "[$(Get-Date -f T)] Converting data from the following BLG files into DAILY FILTERED CSV:"
        $BlgsToConvert | ForEach-Object {$_.Name}
        $i = 0
        foreach ($BlgFile in $BlgsToConvert)
        {
            $i++
            $CsvFile = "$LogDir\tmp\$($BlgFile.BaseName).csv"
            "relog.exe $($BlgFile.FullName) -cf $CounterFile -f csv -o $CsvFile -b $LogBeginTS -e $LogEndTS -y -t $DesampleRatio"
            relog.exe $BlgFile.FullName -cf $CounterFile -f csv -o $CsvFile -b $LogBeginTS -e $LogEndTS -y -t $DesampleRatio
            
        }
        "[$(Get-Date -f T)] Finished Converting $i BLG files into CSV"

        # Zip it.
        "[$(Get-Date -f T)] Compressing log and configuration data to $DailyZipFile"
        Zip-Folder -FolderPath "$LogDir\tmp" -ZipFilePath $DailyZipFile -Overwrite -Verbose | Format-Table -AutoSize
        $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Created DAILY FILTERED ZIP file"
        #endregion dailycpaczip
     }
     #endregion dailyprocessing


    "# -------------------------"
    "[$(Get-Date -f T)] Removing any temporary files..."
    Remove-Item "$LogDir\tmp\*" -Recurse -Force -ErrorAction Ignore
    $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Removed Temporary files"

    #region upload
    "# ----------------"
    "# Data Upload"
    # Random wait seconds (to offset spec updates performed by multiple machines) before uploading
    $UploadEnabled = $RuntimeOptions.UploadEnabled
    $UploadDelayEnabled = $RuntimeOptions.UploadDelayEnabled
    $UploadDelaySecs = $RuntimeOptions.UploadDelaySecs
    $UploadUri = $RuntimeOptions.UploadServiceUri
    $UploadHourly = $RuntimeOptions.UploadToCpacHourly
    If ($UploadDelayEnabled -and ($UploadDelaySecs -gt 0)) {
        $WaitMs = Get-Random -Minimum 0 -Maximum ($UploadDelaySecs * 1000)
        "[$(Get-Date -f T)] Sleeping for a randomized 1-$UploadDelaySecs sec window ($WaitMs ms)..."
        Start-Sleep -Milliseconds $WaitMs
        $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Delayed Upload for $WaitMs ms."
    } else {
        "[$(Get-Date -f T)] No random wait configured; Continuing immediately..."
    }

    # Upload to CPAC
    $UploadedZipSizeMiB = 0.0
    If ($UploadEnabled) {
        If ($UploadHourly) {
            If (Test-Path $HourlyZipFile) {
                $UploadedZipSizeMiB = [math]::Round((([double](Get-Item $HourlyZipFile).Length) / 1MB), 3)
                "[$(Get-Date -f T)] Uploading HOURLY file '$HourlyZipFile' to CPAC..."
                "Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\UploadToCpac.ps1 -UploadFile $HourlyZipFile -UploadServiceURI $UploadUri 2>&1"
                Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\UploadToCpac.ps1 -UploadFile $HourlyZipFile -UploadServiceURI $UploadUri 2>&1
	        "[$(Get-Date -f T)] Upload Finished with code: $LASTEXITCODE"
            $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Uploaded HOURLY ZIP File ($UploadedZipSizeMiB MiB)"
            } Else {
                Write-Error "ERROR: Aborting CPAC HOURLY upload because a '$HourlyZipFile' file does not exist!"
            }
        } Else {
            # Configured to upload once daily; upload a daily file if its the cycle hour.
            If ($HourOfDay -eq $RuntimeOptions.DailyCycleHour) {
                If (Test-Path $DailyZipFile) {
                    $UploadedZipSizeMiB = [math]::Round((([double](Get-Item $DailyZipFile).Length) / 1MB), 3)
                    "[$(Get-Date -f T)] Uploading DAILY file '$DailyZipFile' to CPAC..."
                    "Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\UploadToCpac.ps1 -UploadFile $DailyZipFile -UploadServiceURI $UploadUri 2>&1"
                    Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\UploadToCpac.ps1 -UploadFile $DailyZipFile -UploadServiceURI $UploadUri 2>&1
                    $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Uploaded DAILY ZIP File ($UploadedZipSizeMiB MiB)"
                } Else {
                    Write-Error "ERROR: Aborting CPAC DAILY upload because a '$DailyZipFile' file does not exist!"
                } # End zipfile exists
            } Else {
                "Skipping DAILY upload because current hour ($HourOfDay) is not the daily cycle hour ($($RuntimeOptions.DailyCycleHour))."
            }
        } 
    } Else {
        "Upload is being skipped because UploadEnabled=$UploadEnabled"
    }
    #endregion upload
} # End If $ProcessLogs


#region delayedInventory
"# -------------------------"
"# Delayed (daily) system inventory"
# Run Config dat to gather system inventory data.
# Note, This operation may have a random delay, so perform it last, so as to not further delay data uploads.
If ($PerformDelayedInventory) {
    "[$(Get-Date -f T)] Performing the daily system configuration inventory...."
    "[$(Get-Date -f T)] Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\ConfigJson.ps1"
    "Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\ConfigJson.ps1"
    Powershell -ExecutionPolicy Bypass -NoProfile -File $ScriptDir\ConfigJson.ps1
    $TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Performed DELAYED inventory"
    "[$(Get-Date -f T)] Inventory complete." 
} Else {
    "[$(Get-Date -f T)] Skipping system configuration inventory"
}
#endregion delayedInventory
$TimedOps = Add-Timestamp -TimestampsArray $TimedOps -StepName "Finished Processing"


"# -----------------------------------"
"# SCRIPT FINALIZATION"
$TimedOpsCsv = "$LogDir\TimedOps_$($LastHour.ToString("yyyyMMdd")).CSV"
If (Test-Path $TimedOpsCsv) {
    $TimedOps | Export-CSV "$TimedOpsCsv" -NoTypeInformation -Append -Force -Encoding ASCII
} Else {
    $TimedOps | Export-CSV "$TimedOpsCsv" -NoTypeInformation -Force -Encoding ASCII
}
"--- Script Operations Timing (Also exported to $TimedOpsCsv): ---"
$TimedOps | Select-Object Timestamp, StepName, SecsDelta | Format-Table -AutoSize

"--- Data Characteristics: ---"
"    Uploaded ZIP Size:  $($UploadedZipSizeMiB.ToString('N3')) MiB)"
#Blg Size
$LogDir = (Get-Item $LogDir).FullName
$Files = @(Get-ChildItem "$LogDir\*.blg" -Exclude $RunningBlgLog -ErrorAction Ignore)
If ($Files) {
    $BlgSizeGB = [math]::round((($Files | Measure-Object  Length -Sum).Sum / 1GB), 3)
} Else {
    $BlgSizeGB = 0
}
"    Total BLG Log Size: $($BlgSizeGB.ToString('N3')) GB"

# Granular archive size
$Files = @(Get-ChildItem "$GranularArchiveDir\*.*" -ErrorAction Ignore)
If ($Files) {
    $ArchiveGranularSizeGB = [math]::round((($Files | Measure-Object  Length -Sum).Sum / 1GB), 3)
} Else {
    $ArchiveGranularSizeGB = 0
}
"    Total Granular Archive Size: $($ArchiveGranularSizeGB.ToString('N3')) GB"

# Filtered Hourly Archive
$Files = @(Get-ChildItem "$FilteredHourlyArchiveDir\*.*" -ErrorAction Ignore)
If ($Files) {
    $ArchiveFilteredHourlySizeGB = [math]::round((($Files | Measure-Object Length -Sum).Sum / 1GB), 3)
} Else {
    $ArchiveFilteredHourlySizeGB = 0
}
"    Total FilteredHourly Archive Size: $($ArchiveFilteredHourlySizeGB.ToString('N3')) GB"

#Filtered Daily size
$Files = @(Get-ChildItem "$FilteredDailyArchiveDir\*.*" -ErrorAction Ignore)
If ($Files) {
    $ArchiveFilteredDailySizeGB = [math]::round((($Files | Measure-Object Length -Sum).Sum / 1GB), 3)
} Else {
    $ArchiveFilteredDailySizeGB = 0
}
"    Total FilteredDaily Archive Size: $($ArchiveFilteredDailySizeGB.ToString('N3')) GB"

#Combined total logs size
$Files = @(Get-ChildItem $LogDir\*.* -Exclude $RunningBlgLog -Recurse -ErrorAction Ignore)
If ($Files) {
    $TotalLogSizeGB = [math]::round((($Files | Measure-Object Length -Sum).Sum / 1GB), 3)
} Else {
    $TotalLogSizeGB = 0
}
"    Total Log Size: $($TotalLogSizeGB.ToString('N3')) GB"

$ScriptFinished = Get-Date
$ScriptElapsed = $ScriptFinished - $ScriptStarted
" ### Script completed in $($ScriptElapsed.TotalSeconds.ToString('N3')) secs with $($ERROR.Count) errors. ###"

"[$(Get-Date -f T)] ### Script Complete! ###"
Stop-Transcript

