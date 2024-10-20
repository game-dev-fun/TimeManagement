
$RGBColors = [System.Collections.Generic.HashSet[string]]::new()
$RGBName = @('Green', 'Blue', 'Brown', 'Indigo', 'Lilac', 'Chocolate', 'Eggplant', 'LuckyPoint', 'Peru', 'SeaGreen', 'DarkViolet', 'DarkSlateBlue', 'DarkTangerine', 'DarkTurquoise' , 'Salmon', 'Violet', 'Wasabi', 'WildWillow', 'Viking', 'Orange', 'Clover', 'DarkWood', 'DarkBrown', 'DarkGreenCopper', 'DarkOliveGreen')
$RGBValue = [System.Collections.ArrayList]::new()
[void]$RGBName.ForEach({ $RGBColors.Add($_)
        $num = Get-Random -Minimum 5 -Maximum 40
        $RGBValue.Add($num)
    })


function Get-RandomColor {
    if ($RGBColors.Count -gt 0) {
        $randomColor = $RGBColors | Get-Random
        [void]$RGBColors.Remove($randomColor)
        return $randomColor
    }
    else {
        return $null  # No colors left
    }
}

function ConvertTo-TargetTime {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$inputString
    )
    if ($inputString -match "^\d+$") {
        # Input is just a number, treat it as hours
        return [TimeSpan]::FromHours([double]$inputString)
    }
    elseif ($inputString -match "^\d+:\d+$") {
        # Input is in the format of hours:minutes
        return [TimeSpan]::ParseExact($inputString, "h\:mm", [System.Globalization.CultureInfo]::InvariantCulture)
    }
}

function ConvertTo-FloatTime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [timespan]$time
    )
    $hours = $time.Hours
    $minutes = $time.Minutes * 0.01
    $timeFloat = $hours + $minutes 
    return [float]$timeFloat
}

function Get-Subject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [String]$name
    )
    if ($name.Split('-').Count -gt 1) {
        $overallName = $name.Split('-')[0]
        $subject = $overallName.TrimEnd(' ')
    }
    else {
        $subject = $name.TrimEnd(' ')
    }
    return $subject
}


function Create-Dashboard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$accumulativeFilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$activityFilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$targetActivityFilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$htmlFilePath
    )
    $accumulativeActivityCSV = Import-Csv -Path $accumulativeFilePath 
    $accumulativeActivity = [System.Collections.ArrayList]::new()
    $totalActivity = Import-Csv -Path $activityFilePath
    $targetActivityContent = Get-Content -Path $targetActivityFilePath
    $targetActivityJSON = $targetActivityContent | ConvertFrom-Json 
    $targetActivity = [System.Collections.Hashtable]::new()
    $targetActivityJSON.activities.ForEach({
            $targetActivity[$_.Activity] = ConvertTo-TargetTime $_.Hours
        })
    
    $notRecordedTime = [timespan]::Zero
    $overallTimeHT = [System.Collections.Hashtable]::new()
    $accumulativeNameArray = [System.Collections.ArrayList]::new()
    $accumulativeTimeArray = [System.Collections.ArrayList]::new()
    $accumulativeTargetArray = [System.Collections.ArrayList]::new()
    $totalTime = [timespan]::Zero
    $totalMissedTargets = 0
    foreach ($activity in $accumulativeActivityCSV) {
        $accumulativeTimeSpan = [TimeSpan]::ParseExact($activity.AccumulativeTime, "hh\:mm\:ss", $null)
        if ($activity.Activity -ne "ApplicationRunTime") {
            $notRecordedTime += $accumulativeTimeSpan 
            $subject = Get-Subject $activity.Activity
            if ($overallTimeHT.ContainsKey($subject)) {
                $tempTime = $overallTimeHT[$subject]
                $overallTimeHT[$subject] = $tempTime + $accumulativeTimeSpan 
            }
            else {
                $overallTimeHT[$subject] = $accumulativeTimeSpan 
            }
            if ($targetActivity.Contains($activity.Activity)) {
                $targetTime = $targetActivity[$activity.Activity]
            }
            else {
                $targetTime = [timespan]::Zero
            }
            $missedTarget = $accumulativeTimeSpan - $targetTime 
            $missedTarget = ConvertTo-FloatTime -time $missedTarget
            
            [void]$accumulativeActivity.Add([PSCustomObject]@{Subject = $subject; Activity = $activity.Activity; AccumulativeTime = $accumulativeTimeSpan; TargetTime = $targetTime; TimeOffset = $missedTarget })
            if ($activity.Activity -ne "Essential" -and $activity.Activity -ne "NotRecorded" -and $activity.Activity -ne "Time Waste" ) {
                [void]$accumulativeNameArray.Add($activity.Activity)
                [void]$accumulativeTimeArray.Add([float]$accumulativeTimeSpan.ToString("hh\.mm"))
                [void]$accumulativeTargetArray.Add([float]$targetTime.ToString(("hh\.mm")))
                if ($missedTarget -lt 0) {
                    $totalMissedTargets++
                }
            }
            $targetActivity.Remove($activity.Activity)
        }
        else {
            $totalTime = $accumulativeTimeSpan 
        }
        
    }
    $notRecordedTime = $totalTime - $notRecordedTime
    [void]$accumulativeActivity.Add([PSCustomObject]@{Subject = "NotRecorded"; Activity = "NotRecorded"; AccumulativeTime = $notRecordedTime; TargetTime = ([timespan]::Zero); TimeOffset = (ConvertTo-FloatTime -time $notRecordedTime) })
    $targetActivity.GetEnumerator() | ForEach-Object {
        if ($_.key -ne "ApplicationRunTime") {
            $accumulativeActivity.Add([PSCustomObject]@{Subject = (Get-Subject $_.key); Activity = $_.key; AccumulativeTime = ([timespan]::Zero); TargetTime = $_.Value; TimeOffset = - (ConvertTo-FloatTime -time $_.Value) })
            $subject = Get-Subject $_.Key
            if (-not $overallTimeHT.ContainsKey($subject)) {
                $overallTimeHT[$subject] = [timespan]::Zero
            }  
            if ($_.key -ne "Essential" -and $_.key -ne "NotRecorded" -and $_.key -ne "Time Waste" ) {
                [void]$accumulativeNameArray.Add($_.key)
                [void]$accumulativeTimeArray.Add(0.0)
                [void]$accumulativeTargetArray.Add([float]$_.Value.ToString(("hh\.mm")))
                if ($_.Value -gt 0) {
                    $totalMissedTargets++
                }
            }
        }
    }
    $overallTimeHT["NotRecorded"] = $notRecordedTime
    $overallNameArray = [System.Collections.ArrayList]::new()
    $overallValueArray = [System.Collections.ArrayList]::new()
    $overallColorArray = [System.Collections.ArrayList]::new()
    $overallColorHT = [System.Collections.Hashtable]::new()
    Dashboard -TitleText "Time Data" {
        Add-HTMLStyle -Placement Header -Css @{
            "body" = @{
                "font-size" = "24px"
            }
        }
        New-HTMLTabPanel -Theme elite {
            Tab -Name "Summary" {
                New-HTMLTabPanel -Theme blocks {
                    Tab -Name "Overview" {
                        Section -Height 650 -Invisible {
                            Panel {
                                Chart -Title "Total Time: $($totalTime.ToString("hh\.mm"))" -TitleAlignment center -TitleFontSize 24 -Height 600 {
                                    foreach ($key in $overallTimeHT.Keys) {
                                        [void]$overallNameArray.Add($key)
                                        $tempAccTime = [float]$overallTimeHT[$key].ToString("hh\.mm") 
                                        [void]$overallValueArray.Add($tempAccTime)
                                        $color
                                        if ($key -eq "Time Waste") {
                                            $color = "Red"
                                        }
                                        elseif ($key -eq "NotRecorded") {
                                            $color = "DarkGray"
                                        }
                                        else {
                                            $color = Get-RandomColor
                                            if ($null -eq $color) {
                                                $color = "Black"
                                            }
                                        }
                                        [void]$overallColorArray.Add($color)
                                        $overallColorHT[$key] = $color
                                        New-ChartPie -Name $key -Value $tempAccTime -Color  $color
                                    }
                                    New-ChartEvent -DataTableID "accumulativeTable" -ColumnID 0
                                }
                            }
                
                            Panel {
                                Chart -Title "Total Time: $($totalTime.ToString("hh\.mm"))" -TitleAlignment center -TitleFontSize "24px" -Height 600 {
                                    New-ChartLegend -LegendPosition bottom -HorizontalAlign right -Color $overallColorArray
                                    New-ChartAxisY -LabelMaxWidth 200 -LabelAlign left -Show -TitleText 'Subject' -TitleColor Red
                                    New-ChartBarOptions -Distributed
                                    for ($i = 0; $i -lt $overallNameArray.Count; $i++) {
                                        New-ChartBar -Name $overallNameArray[$i] -Value $overallValueArray[$i]
                                    }
                                    New-ChartEvent -DataTableID "accumulativeTable" -ColumnID 0
                                }
                            }
            
                        }

                    }
                    Tab -Name "Detailed" {
                        Section -Height 650 -Invisible {
                            foreach ($name in $overallNameArray) {
                                Chart -Title "$name : $($overallTimeHT[$name].ToString("hh\.mm"))" -TitleAlignment center -TitleFontSize "18px" -Height 600 {
                                    ChartBarOptions -Type barStacked -Vertical
                                    New-ChartToolbar pan
                                    $tempActivity = $accumulativeActivity | where Subject -eq $name
                                    $activityNameArray = [System.Collections.ArrayList]::new()
                                    $activityValueArray = [System.Collections.ArrayList]::new()
                                    $tempActivity | ForEach-Object {
                                        if ($_.Activity.Split('-').Count -gt 1) {
                                            $activityName = $_.Activity.Split('-')[1]
                                        }
                                        else {
                                            $activityName = $_.Activity
                                        }
                                        [void]$activityNameArray.Add($activityName.TrimStart(' '))
                                        $tempAccTime = [TimeSpan]::ParseExact($_.AccumulativeTime, "hh\:mm\:ss", $null)
                                        [void]$activityValueArray.Add([float]$tempAccTime.ToString("hh\.mm"))
                                    }
                                    if ($name -eq "Time Waste") {
                                        ChartLegend -Names $activityNameArray -Color Red
                                    }
                                    else {
                                        ChartLegend -Names $activityNameArray 
                                    }
                                    ChartBar -Name $name -Value $activityValueArray
                                    New-ChartEvent -DataTableID "accumulativeTable" -ColumnID 0
                                }
                            }
                        }
                    }
                    Tab -Name "Missed Targets" {
                        Section -Height 650 -Invisible {
                            Chart -Title "Missed Targets: $totalMissedTargets" -TitleAlignment center -TitleFontSize "24px" -Height 600 {
                                ChartAxisX -Names $accumulativeNameArray
                                ChartLine -Name "TimePlanned" $accumulativeTargetArray -Color Red -Dash 4 -Curve smooth
                                ChartLine -Name "TimeSpent" $accumulativeTimeArray -Color Green -Curve smooth -Dash 0
                                New-ChartEvent -DataTableID "accumulativeTable" -ColumnID 1
                            }
                        }
                    }
                    Tab -Name "Target Bar" {
                        Section -Height 650 -Invisible {
                            Panel {
                                Table -DataTable $overallNameArray -DataTableID "overallTable" -HideButtons -HideFooter {
                                    New-TableEvent -TableID "accumulativeTable" -SourceColumnName "Name" -TargetColumnID 0 -SourceColumnID 0
                                }
                                
                            }
                            Panel {
                                foreach ($name in $overallNameArray) {
                                    "<div data-name = '$name' style='display:none;'>"
                                    Chart -Title "$name : $($overallTimeHT[$name].ToString("hh\.mm"))" -TitleAlignment center -TitleFontSize "18px" -Height 600 {
                                        ChartBarOptions -Type barStacked
                                        ChartLegend -Names @("Time Spent", "Time Offset") -Color @("Green", "Red")
                                        foreach ($act in $accumulativeActivity) {
                                            if ($name -eq $act.Subject) {
                                                $value = [System.Collections.ArrayList]::new()
                                                $value.Add([float]$act.AccumulativeTime.ToString("hh\.mm"))
                                                if ($act.TimeOffset -lt 0) {
                                                    $value.Add([Math]::Abs($act.TimeOffset))
                                                }
                                                else {
                                                    $value.Add(0)
                                                }
                                                ChartBar -Name $act.Activity -Value $value
                                            }
                                        }
                                        #New-ChartEvent -DataTableID "accumulativeTable" -ColumnID 1
                                    }
                                    
                                    "</div>"
                                }
                            }
                        }
                        New-HTMLPanel {
                            "<script>
                            document.addEventListener('DOMContentLoaded', function() {
                                var rows = document.querySelectorAll('#overallTable tr');
                                rows.forEach(function(row) {
                                    row.addEventListener('click', function() {
                                        var selectedName = this.cells[0].textContent.trim(); // Get the first column
                      
                                        // Selecting the chart based on data-name
                                        var chartToShow = document.querySelector('div[data-name=`"' + selectedName + '`"]');
                      
                                        if (chartToShow) {
                                            var allCharts = document.querySelectorAll('div[data-name]');
                                            allCharts.forEach(function(chart) {
                                                chart.style.display = 'none'; // Hiding all charts
                                            });
                                            chartToShow.style.display = 'block'; // Displaying only the selected chart
                                        } else {
                                            console.error('Chart not found: ' + selectedName);
                                        }
                                    });
                                });
                            });
                            </script>"
                        }      
                    }
                }
                Section -HeaderText "Time Log" {
                    Panel {
                        Table -DataTable $accumulativeActivity -DataTableID "accumulativeTable" -HideButtons -HideFooter -PagingLength 8 {
                            New-TableCondition -Operator lt -ComparisonType number -Name "TimeOffset" -Value 0 -BackgroundColor Red -Color White -Row
                            New-TableCondition -Operator ge -ComparisonType number -Name "TimeOffset" -Value 0 -BackgroundColor Green -Color White -Row
                            New-TableConditionGroup -Logic OR -BackgroundColor Gray -Color White -Row -Conditions {        
                                New-TableCondition -Operator eq -ComparisonType string -Name "Activity" -Value "Essential"
                                New-TableCondition -Operator eq -ComparisonType string -Name "Activity" -Value "NotRecorded"
                            }
                            New-TableCondition -Operator eq -ComparisonType string -Name "Activity" -Value "Time Waste" -BackgroundColor Black -Color White -Row
                           
                        } 
                    }
                }
            }
            Tab -Name "Calendar" {
                New-HTMLCalendar -InitialView listMonth -HeaderLeft listDay -HeaderRight listMonth {                
                    foreach ($ttlActivity in $totalActivity) {
                        if ($ttlActivity.Activity -ne "ApplicationRunTime") {                        
                            $subject = Get-Subject $ttlActivity.Activity
                            if ($overallColorHT.ContainsKey($subject)) {
                                $color = $overallColorHT[$subject]
                            }
                            else {
                                $color = Get-RandomColor
                                if ($null -eq $color) {
                                    $color = "Gray"
                                }
                            }
                            New-CalendarEvent -Title $ttlActivity.Activity  -StartDate ([datetime]$ttlActivity.StartTime) -EndDate ([datetime]$ttlActivity.EndTime) -Color $color
                        }
                    }
                }
            }
        }
    } -ShowHTML -FilePath $htmlFilePath
    Remove-Variable accumulativeActivityCSV -ErrorAction Ignore
    Remove-Variable accumulativeActivity -ErrorAction Ignore
    Remove-Variable totalActivity -ErrorAction Ignore
    Remove-Variable targetActivityContent -ErrorAction Ignore
    Remove-Variable targetActivityJSON -ErrorAction Ignore
    Remove-Variable targetActivity -ErrorAction Ignore
    Remove-Variable notRecordedTime -ErrorAction Ignore
    Remove-Variable overallTimeHT -ErrorAction Ignore
    Remove-Variable accumulativeNameArray -ErrorAction Ignore
    Remove-Variable accumulativeTimeArray -ErrorAction Ignore
    Remove-Variable accumulativeTargetArray -ErrorAction Ignore
    Remove-Variable totalTime -ErrorAction Ignore
    Remove-Variable overallTimeHT -ErrorAction Ignore
    Remove-Variable overallNameArray -ErrorAction Ignore
    Remove-Variable overallValueArray -ErrorAction Ignore
    Remove-Variable overallColorArray -ErrorAction Ignore
}          
   


# $accumulativeFilePath = "C:\Users\shivam madaan\Documents\Scripts\TimeManagement\Time\October\2\Time\2-October-Accumulative.csv"
# $totalTimePath = "C:\Users\shivam madaan\Documents\Scripts\TimeManagement\Time\October\2\Time\2-October-Time.csv"
#$targetActivityFilePath = "C:\Users\shivam madaan\Documents\Scripts\TimeManagement\Time\October\2\Config\2-October-Config.json"
# $htmlFilePath = "C:\Users\shivam madaan\Documents\Scripts\TimeManagement\Time\October\2\Report\2-October-Report.html"
# Create-Dashboard -accumulativeFilePath $accumulativeFilePath -activityFilePath $totalTimePath -targetActivityFilePath $configFilePath -htmlFilePath $htmlFilePath

 