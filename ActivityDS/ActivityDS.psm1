Remove-Variable Activities -ErrorAction Ignore
Remove-Variable currentActivity -ErrorAction Ignore
Remove-Variable possibleActivities -ErrorAction Ignore
Remove-Variable activityFilePath -ErrorAction Ignore

$global:Activities = [System.Collections.Hashtable]::new()
$global:currentActivity = $null
$global:possibleActivities = [System.Collections.Generic.HashSet[System.String]]::new()
$global:activityFilePath = $null
$global:accumulativeFilePath = $null

function Set-ActivityFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$csvActivityFilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$csvAccumulativeFilePath
    )
    $global:activityFilePath = $csvActivityFilePath
    if (-not (Test-Path $global:activityFilePath)) {
        $csvContent = @("Activity,StartTime,EndTime,TotalTime,AccumulativeTime")
        $csvContent | Out-File -FilePath $global:activityFilePath -Encoding utf8 -Force
    }
    else {
        Write-Verbose "Activity File Path exist: $($global:activityFilePath)"
    }

    $global:accumulativeFilePath = $csvAccumulativeFilePath 
    if (Test-Path $global:accumulativeFilePath) {
        Load-Time
        Write-Verbose "Accumulative File Path exist: $($global:accumulativeFilePath)"
    }
    else {
        Write-Verbose "No Accumulative File found"
    }
}

class TimeBlock {
    $StartTime
    $EndTime
    $TotalTime
    $AccumulativeTime

    TimeBlock([datetime]$st, [datetime]$et, [timespan]$tt, [timespan]$at) {
        [datetime]$this.StartTime = $st
        [datetime]$this.EndTime = $et
        [timespan]$this.TotalTime = $tt
        [timespan]$this.AccumulativeTime = $at 
    }
    
    TimeBlock([timespan]$previousTime) {
        [datetime]$this.StartTime = [datetime]::now
        $this.EndTime = $null
        $this.TotalTime = $null
        [timespan]$this.AccumulativeTime = $previousTime
    }

    # Set & calculate timespan
    SetEndTime() {
        [datetime]$this.EndTime = [datetime]::now
        [timespan]$this.TotalTime = $this.EndTime - $this.StartTime
        [timespan]$this.AccumulativeTime += $this.TotalTime
    }
    
    # returns a CSV string of the member
    [string] PrintCSV() {
        $stringStartTime = $this.StartTime.ToString("MM/dd/yyyy HH:mm:ss")
        if ($null -eq $this.EndTime ) {
            $tempEndTime = [datetime]::now
            $stringEndTime = $tempEndTime.ToString("MM/dd/yyyy HH:mm:ss")
    
            $tempTotalTime = $tempEndTime - $this.StartTime
            $stringTotalTime = $tempTotalTime.ToString("hh\:mm\:ss")
            $stringAccumulativeTime = ($tempTotalTime + $this.AccumulativeTime).ToString("hh\:mm\:ss")
        }
        else {
            $stringEndTime = $this.EndTime.ToString("MM/dd/yyyy HH:mm:ss")
            $stringTotalTime = $this.TotalTime.ToString("hh\:mm\:ss")
            $stringAccumulativeTime = $this.AccumulativeTime.ToString("hh\:mm\:ss")
        }
        return "$stringStartTime,$stringEndTime,$stringTotalTime,$stringAccumulativeTime"
    }
}

#check if Activity Exists in Possible Activities
function Check-Activity {
    [Cmdletbinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Activity,
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$appRun = "ApplicationRunTime"
    )
    return $Activity -eq $appRun -or $global:possibleActivities.Contains($Activity)
}

#Add Activity to Possible Activities
function Add-Activity {
    [Cmdletbinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Activity
    )
    if (-not (Check-Activity $Activity)) {
        [void]$global:possibleActivities.Add($Activity)
    }
    else {
        Write-Verbose "$Activity already exist"
    }
}



#Start Time for Activity
function Start-Time {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Activity,
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$appRun = "ApplicationRunTime"
    )
    if (Check-Activity $Activity) {
        if($Activity -eq $appRun -and $global:Activities.ContainsKey($Activity) -and $null -eq $global:Activities[$Activity].EndTime){
                Write-Verbose "$Activity is already active"
                return
        }
        elseif ($global:currentActivity -eq $Activity) {
            Write-Verbose "$Activity is already active"
            return
        }
        elseif ($Activity -ne $appRun -and $null -ne $global:currentActivity) {
            Write-Verbose "Ending $global:currentActivity which is active."
            End-Time $global:currentActivity
        }
        
        if ($global:Activities.ContainsKey($Activity)) {
            $previousTime = $global:Activities[$Activity].AccumulativeTime 
            $global:Activities[$Activity] = [TimeBlock]::new($previousTime)
        }
        else {
            $global:Activities[$Activity] = [TimeBlock]::new([timespan]0)
        }

        if ($Activity -ne $appRun) {
            $global:currentActivity = $Activity
        }
        Write-Verbose "$Activity is now active"
    }
    else {
        Write-Verbose "$Activity not found"
    }
}

#check if the activity is active

function Check-ActiveActivity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Activity,
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$appRun = "ApplicationRunTime"
    )
    if ($Activity -ne $appRun -and $global:currentActivity -ne $Activity) {
        Write-Verbose "$Activity is not currently Active"
        return $false
    }
    elseif ($global:Activities.ContainsKey($Activity)) {
        if ($null -eq $global:Activities[$Activity]) {
            Write-Verbose "No $Activity active"
            return $false
        }
        else {
            $actTimeBlock = $global:Activities[$Activity]
            if ($null -eq $actTimeBlock.StartTime -or $null -ne $actTimeBlock.EndTime) {
                Write-Verbose "No $Activity active"
                return $false
            }
            else {
                return $true
            }
        }
    }
    return $false
}


#end the timer for activity and print it to the activity file path
function End-Time {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Activity,
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$appRun = "ApplicationRunTime"
    )

    if (Check-ActiveActivity $Activity) {
        $global:Activities[$Activity].SetEndTime()
        "$Activity,$($global:Activities[$Activity].printCSV())" | Out-File -FilePath $global:activityFilePath -Append -Encoding utf8
        if($Activity -ne $appRun){
            $global:currentActivity = $null
        }
    }
    else {
        Write-Verbose "$Activity not active"
    }
    Save-Activities
}


function Save-Activities {

    # Create header row
    $csvContent = @("Activity,StartTime,EndTime,TotalTime,AccumulativeTime")
    foreach ($key in $global:Activities.Keys) {
        $csvContent += "$key,$($global:Activities[$key].PrintCSV())"
    }
    $csvContent | Out-File -FilePath $global:accumulativeFilePath -Encoding UTF8 -Force
}


function Get-CurrentActivity {
    return $global:currentActivity
}


function Get-AccumulativeTime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Activity
    )
    if ($global:Activities.ContainsKey($Activity)) {
        if ($null -eq $global:Activities[$Activity]) {
            [timespan]$zero = [timespan]0
            return $zero
        }
        else {
            return $global:Activities[$Activity].AccumulativeTime
        }
    }
    else {
        [timespan]$zero = [timespan]0
        return $zero
    }
}


function Load-Time{
    if(Test-Path $global:accumulativeFilePath){
        $global:Activities.Clear()
        $contentCSV = Import-CSV -Path $global:accumulativeFilePath
        $contentCSV | ForEach-Object{
            $startTime = [datetime]::ParseExact($_.StartTime, "MM/dd/yyyy HH:mm:ss", $null)
            $endTime = [datetime]::ParseExact($_.EndTime, "MM/dd/yyyy HH:mm:ss", $null)
            $totalTime = [TimeSpan]::ParseExact($_.TotalTime, "hh\:mm\:ss", $null)
            $accumulativeTime = [TimeSpan]::ParseExact($_.AccumulativeTime, "hh\:mm\:ss", $null)
            $global:Activities[$_.Activity] = [TimeBlock]::new($startTime, $endTime, $totalTime, $accumulativeTime)
        }
        
    }
}



#load detailed time from time csv file
function Load-DetailTime {
    if (Test-Path $global:activityFilePath) {
        $contentCSV = Import-Csv -Path $global:activityFilePath
        $contentCSV | ForEach-Object {
            $startTime = [datetime]::ParseExact($_.StartTime, "MM/dd/yyyy HH:mm:ss", $null)
            $endTime = [datetime]::ParseExact($_.EndTime, "MM/dd/yyyy HH:mm:ss", $null)
            $totalTime = [TimeSpan]::ParseExact($_.TotalTime, "hh\:mm\:ss", $null)
            $accumulativeTime = [TimeSpan]::ParseExact($_.AccumulativeTime, "hh\:mm\:ss", $null)
            if ($global:Activities.ContainsKey($_.Activity)) {
                if ($null -eq $global:Activities[$_.Activity]) {
                    $global:Activities[$_.Activity] = [System.Collections.ArrayList]::new()
                }
                $global:Activities[$_.Activity].Add([TimeBlock]::new($startTime, $endTime, $totalTime, $accumulativeTime))
            }
            else {
                $global:Activities[$_.Activity] = [System.Collections.ArrayList]::new()
                $global:Activities[$_.Activity].Add([TimeBlock]::new($startTime, $endTime, $totalTime, $accumulativeTime))
            }
        }
    }
}

function Get-StartTime{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Activity
    )
    if (Check-ActiveActivity $Activity) {
       return $global:Activities[$Activity].StartTime
    }
    else {
        return $null
    }
}
