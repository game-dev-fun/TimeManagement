$status = $true
while ($global:status) {

    $ActivitySettingsModule = Resolve-Path -LiteralPath "$PSScriptRoot\ActivitySettings"
    Import-Module $ActivitySettingsModule -Force
    $ActivityDSModule = Resolve-Path -LiteralPath "$PSScriptRoot\ActivityDS"
    Import-Module $ActivityDSModule -Force
    $ChartModule = Resolve-Path -LiteralPath "$PSScriptRoot\Chart"
    Import-Module $ChartModule -Force 

    #convert from string to [timespan]
    function ConvertTo-TimeSpan {
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

    #get templated config files
    function getConfigFile {
        $scheduleFolderPath = Join-Path -Path $global:scriptFolder -ChildPath "Schedule" 
        $defaultConfigPath = Join-Path -Path $scheduleFolderPath -ChildPath "Default.json"
        $todayPath = Join-Path -Path $scheduleFolderPath -ChildPath "$([datetime]::Now.DayOfWeek.ToString()).json"
        if (Test-Path -Path $todayPath) {
            Write-Verbose "$([datetime]::Now.DayOfWeek.ToString()).json template config file selected"
            return $todayPath
        }
        else {
            if (Test-Path -Path $defaultConfigPath) {
                Write-Verbose "Default.config template file selected"
                return $defaultConfigPath
            }
            else {
                Write-Verbose "No Template Config file found"
                return $null
            }
        }
    }
    #initialize the folders
    try {
        $scriptFolder = $PSScriptRoot
        $settingPath = Join-Path -Path $scriptFolder -ChildPath "Config.json"
        $settingContent = Get-Content -Path $settingPath
        $settingJSON = $settingContent | ConvertFrom-Json
        $activityFolderPath = $settingJSON.activityFolder
        $month = Get-Date -Format "MMMM"
        $monthFolderPath = Join-Path -Path $activityFolderPath -ChildPath $month
        $todayConfigFilePath = Join-Path -Path $monthFolderPath -ChildPath "$([datetime]::Now.Day)\Config\$([datetime]::Now.Day)-$month-Config.json"
        $todayConfigBackupFilePath = Join-Path -Path $monthFolderPath -ChildPath "$([datetime]::Now.Day)\Config\$([datetime]::Now.Day)-$month-Config-Backup.json"
        $todayConfigFolderPath = Split-Path $todayConfigFilePath 
        $todayTimeFolderPath = Join-Path -Path $monthFolderPath -ChildPath "$([datetime]::Now.Day)\Time\"
        $todayTimeFilePath = Join-Path -Path $todayTimeFolderPath -ChildPath "$([datetime]::Now.Day)-$month-Time.csv"
        $todayAccmulativeFilePath = Join-Path -Path $todayTimeFolderPath -ChildPath "$([datetime]::Now.Day)-$month-Accumulative.csv"
        $todayReportFolderPath = Join-Path -Path $monthFolderPath -ChildPath "$([datetime]::Now.Day)\Report\"
        $todayReportFilePath = Join-Path -Path $todayReportFolderPath -ChildPath "$([datetime]::Now.Day)-$month-Report.html"
        $initialize = $false
        if (-not (Test-Path $todayConfigFilePath)) {
            Write-Verbose "todayConfigFile not found"
            $initialize = $true
        }
        if (-not(Test-Path $todayConfigFolderPath)) {
            New-Item -ItemType Directory -Path $todayConfigFolderPath -Force
        }
        if (-not(Test-Path $todayTimeFolderPath)) {
            New-Item -ItemType Directory -Path $todayTimeFolderPath -Force
        }
        if (-not(Test-Path $todayReportFolderPath)) {
            New-Item -ItemType Directory -Path $todayReportFolderPath -Force
        }

        $configTemplatePath = getConfigFile
        while ($initialize) {
            Show-Setting -configPath $configTemplatePath -todayConfigPath $todayConfigFilePath -initialize $true -backupConfigPath $todayConfigBackupFilePath
            $initialize = -not $(Test-Path -Path $todayConfigFilePath)
        }
        Set-ActivityFile -csvActivityFilePath $todayTimeFilePath -csvAccumulativeFilePath $todayAccmulativeFilePath
    }
    catch {
        Write-Debug "Initialization Error: $($_.Exception.Message)" 
    }

    #setting activity timer
    function Set-LabelTimer {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [System.Windows.Forms.ToolStripMenuItem]$menu
        )
        #resetting timer when new Activity
        if (Check-ActiveActivity $menu.Activity) {
            $btnTimerActivity.Text = $menu.Activity
            $btnTimerActivity.Menu = $menu
            Set-ActivityTimer -active $true
            $lblSessionTimer.Text = "00:00"
            Set-LabelSize $lblSessionTimer
            $lblTotalTimer.Activity = $menu.Activity
            $lblTotalTimer.PreviousTime = ([TimeSpan]::Zero)
            $lblTotalTimer.TargetTime = $menu.TargetTime
            $lblTotalTimer.Text = "($($menu.AccumulatedTime.ToString("hh\:mm"))/$($lblTotalTimer.TargetTime.ToString("hh\:mm")))"
            if ($menu.Activity -ne "Time Waste" -and $menu.Activity -ne "Essential") {
                if ($menu.AccumulatedTime -ge $lblTotalTimer.TargetTime) {
                    $lblTotalTimer.ForeColor = "Green"
                }
                else {
                    $lblTotalTimer.ForeColor = "Red"
                }
            }
            else {
                if ($menu.AccumulatedTime -gt $lblTotalTimer.TargetTime) {
                    $lblTotalTimer.ForeColor = "Red"
                }
                else {
                    $lblTotalTimer.ForeColor = "Green"
                }
            }
        
            Set-LabelSize $lblTotalTimer
            $timer.start()
        }
        #if ending timer
        else {
            if ($timer.Enabled) {
                $timer.Stop()
            }
            $btnTimerActivity.Text = "No Active Timer"
            Set-ActivityTimer -active $false
            $btnTimerActivity.Menu = $null
            $lblSessionTimer.Text = ""
            $lblSessionTimer.Size = [System.Drawing.Size]::new(0, 0)
            $lblTotalTimer.Text = ""
            $lblTotalTimer.Size = [System.Drawing.Size]::new(0, 0)
            $lblTotalTimer.Activity = $null
            $lblTotalTimer.TargetTime = ([TimeSpan]::Zero)

        }
        #update timer label location according to the text
        Set-ButtonSize $btnTimerActivity
        $lblSessionTimerX = $btnTimerActivity.Size.Width + $btnTimerActivity.Location.X + $formPaddingX
        $lblSessionTimer.Location = [System.Drawing.Point]::new($lblSessionTimerX, $formPaddingY)
        $lblTotalTimerX = $lblSessionTimer.Size.Width + $lblSessionTimer.Location.X + $formPaddingX
        $lblTotalTimer.Location = [System.Drawing.Point]::new($lblTotalTimerX, $formPaddingY)
        $btnActivityX = $lblTotalTimer.Size.Width + $lblTotalTimer.Location.X + $formPaddingX
        $btnActivity.Location = New-Object System.Drawing.Point($btnActivityX, $formPaddingY)
        $btnEssentialLocationX = $btnActivity.Location.X + $btnActivity.Width + $btnPaddingX
        $btnEssential.Location = [System.Drawing.Point]::new($btnEssentialLocationX, $formPaddingY)
        $btnChartLocationX = $btnEssential.Location.X + $btnEssential.Width + $btnPaddingX
        $btnChart.Location = [System.Drawing.Point]::new($btnChartLocationX, $formPaddingY)
        $btnTimeTargetLocationX = $btnChart.Location.X + $btnChart.Width + $btnPaddingX
        $btnTimeTarget.Location = [System.Drawing.Point]::new($btnTimeTargetLocationX, $formPaddingY)
    }

    #break aka essential timer
    function Set-EssentialTimer {
        #when break button is pressed
        if (-not $btnEssential.Pressed) {
            #if previous activity is running save it's elapsed time 
            if ($timer.Enabled) {
                $lblTotalTimer.PreviousTime = $lblTotalTimer.ElapsedTime
                $timer.Stop()
            }
            Set-ActivityButton -active $false
            Set-ActivityTimer -active $false
            $btnEssential.Pressed = $true
            $btnEssential.PreviousActivity = Get-CurrentActivity
            Start-Time "Essential" 
            $btnEssential.add_Paint({
                    param($sender, $e)
                    $btn = [System.Windows.Forms.Button]$sender
                    [System.Windows.Forms.ControlPaint]::DrawBorder(
                        $e.Graphics, 
                        $btn.ClientRectangle,
                        [System.Drawing.Color]::Red, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Red, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Maroon, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Maroon, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                    )
                })
            $btnEssential.Text = "00:00:00"
            $btnEssential.Stopwatch.Restart()
            $btnEssential.StopWatch.Start()
            $btnEssential.Timer.Start()
            $btnEssential.ForeColor = "Red"
        }
        #when break button is released
        else {
            Set-ActivityButton -active $true
            $btnEssential.Pressed = $false
            if ($null -ne $btnEssential.PreviousActivity) {
                Start-Time $btnEssential.PreviousActivity
                $timer.Start()
            }
            else {
                End-Time "Essential" 
            }
            Set-ActivityTimer -active $true
            $btnEssential.Stopwatch.Stop()
            $btnEssential.Timer.Stop()
            Set-ActivityMenu -activity "Essential" -getAccumulativeTime $true
            $btnEssential.Text = "Break"
            $btnEssential.ForeColor = "LightSeaGreen"
            $btnEssential.add_Paint({
                    param($sender, $e)
                    $btn = [System.Windows.Forms.Button]$sender
                    [System.Windows.Forms.ControlPaint]::DrawBorder(
                        $e.Graphics, 
                        $btn.ClientRectangle,
                        [System.Drawing.Color]::GreenYellow, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::GreenYellow, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::LightSeaGreen, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::LightSeaGreen, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                    )
                })
        }
    }

    function Set-ActivityMenu {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [string]$activity,
            [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [bool]$getAccumulativeTime = $false,
            [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [timespan]$accumulativeTimeValue = [timespan]::Zero
        )
        if ($menuIndexHT.ContainsKey($activity)) {
            $i = $menuIndexHT[$activity]
            if ($getAccumulativeTime) {
                $contextMenu.Items[$i].AccumulatedTime = Get-AccumulativeTime -Activity $activity
            }
            else {
                $contextMenu.Items[$i].AccumulatedTime = $accumulativeTimeValue 
            }
            $accTime = $contextMenu.Items[$i].AccumulatedTime
            $targetTime = $contextMenu.Items[$i].TargetTime
            $contextMenu.Items[$i].Text = $activity + " | $($accTime.ToString("hh\:mm")) / $($targetTime.ToString("hh\:mm"))"
            if ($activity -ne "Time Waste" -and $activity -ne "Essential") {
                if ($accTime -ge $targetTime) {
                    $contextMenu.Items[$i].ForeColor = "Green"
                }
                else {
                    $contextMenu.Items[$i].ForeColor = "Red"
                }
            }
            else {
                if ($accTime -gt $targetTime) {
                    $contextMenu.Items[$i].ForeColor = "Red"
                }
                else {
                    $contextMenu.Items[$i].ForeColor = "Green"
                }
            }
        
        }
    }

    #Timer message box
    function showActivityTime {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [System.Windows.Forms.ToolStripMenuItem]$menu
        )
        if ($menu.Activity -eq (Get-CurrentActivity)) {
            if ($lblTotalTimer.Activity -eq $menu.Activity) {
                $tempTime = $lblTotalTimer.Text
            }
            else {
                $tempTime = $menu.accumulatedTime.ToString("hh\:mm\:ss")
            }
            $timeMsg = "`nTotal Time: $tempTime`nTarget Time: $($menu.targetTime.ToString("hh\:mm"))`n"
            $msg = "Do you want to end $($menu.Activity)" + $timeMsg
        }
        else {
            $timeMsg = "`nTotal Time: $($menu.accumulatedTime.ToString("hh\:mm\:ss"))`nTarget Time: $($menu.targetTime.ToString("hh\:mm"))`n"
            $msg = "Do you want to start $($menu.Activity)" + $timeMsg
        }
        $result = [System.Windows.Forms.MessageBox]::Show(
            "$msg",
            "Activity timer", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question
        )   


        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            if ($timer.Enabled) {
                $timer.Stop()
            }
            #if current activity is already running then ask users if they want to end it
            $previousActivity = Get-CurrentActivity
            if ((Get-CurrentActivity) -eq $menu.Activity) {
                End-Time $menu.Activity 
            }
            #if some other activity is running then start this activity
            elseif ((Get-CurrentActivity) -ne $menu.Activity) {
                Start-Time $menu.Activity 
            }
            if ($null -ne $previousActivity) {
                Set-ActivityMenu -activity $previousActivity -getAccumulativeTime $true
            }
            Set-LabelTimer $menu
        }

 
    }

    #set label size dynamically according to the text width
    function Set-LabelSize {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [System.Windows.Forms.Label]$labelTemp
        )
        $graphics = $labelTemp.CreateGraphics()
        $tempText = $labelTemp.Text + ':'
        $textSize = $graphics.MeasureString($tempText, $labelTemp.Font)
        $labelTemp.Size = [System.Drawing.Size]::new($textSize.Width, $textSize.Height)
        $graphics.dispose()
    }

    #set button size dynamically according to it's text width
    function Set-ButtonSize {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [System.Windows.Forms.Button]$labelTemp
        )
        $graphics = $labelTemp.CreateGraphics()
        $tempText = $labelTemp.Text + ':'
        $textSize = $graphics.MeasureString($tempText, $labelTemp.Font)
        $labelTemp.Size = [System.Drawing.Size]::new($textSize.Width, $textSize.Height)
        $graphics.dispose()
    }

    function Set-ActivityButton {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [bool]$active
        )
        if ($active) {
            $btnActivity.Enabled = $true
            $btnActivity.add_Paint({
                    param($sender, $e)
                    $btn = [System.Windows.Forms.Button]$sender
                    [System.Windows.Forms.ControlPaint]::DrawBorder(
                        $e.Graphics, 
                        $btn.ClientRectangle,
                        [System.Drawing.Color]::Pink, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Pink, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::OrangeRed, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::OrangeRed, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                    )
                })
        }
        else {
            $btnActivity.Enabled = $false
            $btnActivity.add_Paint({
                    param($sender, $e)
                    $btn = [System.Windows.Forms.Button]$sender
                    [System.Windows.Forms.ControlPaint]::DrawBorder(
                        $e.Graphics, 
                        $btn.ClientRectangle,
                        [System.Drawing.Color]::LightGray, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::LightGray, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Black, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Black, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                    )
                })
        }
        $btnActivity.Invalidate()
    }



    function Set-ActivityTimer {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [bool]$active
        )

        if ($active -and ($null -ne (Get-CurrentActivity))) {
            $btnTimerActivity.Enabled = $true
            $btnTimerActivity.add_Paint({
                    param($sender, $e)
                    $btn = [System.Windows.Forms.Button]$sender
                    [System.Windows.Forms.ControlPaint]::DrawBorder(
                        $e.Graphics, 
                        $btn.ClientRectangle,
                        [System.Drawing.Color]::Pink, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Pink, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::OrangeRed, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::OrangeRed, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                    )
                })
        }
        else {
            $btnTimerActivity.Enabled = $false
            $btnTimerActivity.add_Paint({
                    param($sender, $e)
                    $btn = [System.Windows.Forms.Button]$sender
                    [System.Windows.Forms.ControlPaint]::DrawBorder(
                        $e.Graphics, 
                        $btn.ClientRectangle,
                        [System.Drawing.Color]::LightGray, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::LightGray, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Black, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                        [System.Drawing.Color]::Black, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                    )
                })
        }
        $btnTimerActivity.Invalidate()
    }


    try {
    
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    
        Start-Time "ApplicationRunTime" 
   
        $screen = $global:settingJSON.screenLocation
    
        $form = [System.Windows.Forms.Form]::new()
    
        $form.TopMost = $true
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual

        $workingArea = [System.Windows.Forms.Screen]::AllScreens[$screen].WorkingArea
        $totalArea = [System.Windows.Forms.Screen]::AllScreens[$screen].Bounds
        $form.Location = [System.Drawing.Point]::new($workingArea.Left, $workingArea.Bottom)
        $formWidth = $workingArea.Right - $workingArea.Left
        $formHeight = $totalArea.Height - $workingArea.Height
        $form.ClientSize = [System.Drawing.Size]::new($formWidth, $formHeight)
    
        $form.BackColor = [System.Drawing.Color]::FromArgb(48, 48, 48)
        $defaultFont = [System.Drawing.Font]::new('Arial', 24)
        $form.Font = $defaultFont
        $menuFont = New-Object System.Drawing.Font("Arial", 20)  # Change size to 12, you can adjust this

        $todayActivitiesFile = Get-Content -Path $todayConfigFilePath
        $todayActivitiesJson = $todayActivitiesFile | ConvertFrom-Json
        $todayActivities = $todayActivitiesJson.Activities
        $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

        $menuIndexHT = [System.Collections.Hashtable]::new()
        $index = 0
        foreach ($todayActivity in $todayActivities) {
            $menuActivity = $todayActivity.Activity
            Add-Activity $menuActivity
            ${$menuActivity} = New-Object System.Windows.Forms.ToolStripMenuItem
            Add-Member -InputObject ${$menuActivity} -MemberType NoteProperty -Name "Activity" -Value $menuActivity
            [timespan]$accTime = Get-AccumulativeTime -Activity $menuActivity
            Add-Member -InputObject ${$menuActivity} -MemberType NoteProperty -Name "AccumulatedTime" -Value $accTime
            if (-not [string]::IsNullOrEmpty($todayActivity.hours)) {
                [timespan]$requiredTime = ConvertTo-TimeSpan $todayActivity.Hours
            }
            else {
                [timespan]$requiredTime = ([TimeSpan]::Zero)
            }
            Add-Member -InputObject ${$menuActivity} -MemberType NoteProperty -Name "TargetTime" -Value $requiredTime
            ${$menuActivity}.Text = $menuActivity + " | $($accTime.ToString("hh\:mm")) / $($requiredTime.ToString("hh\:mm"))"
            ${$menuActivity}.Font = $menuFont
            if ($menuActivity -ne "Time Waste" -and $menuActivity -ne "Essential") {
                if ($accTime -ge $requiredTime) {
                    ${$menuActivity}.ForeColor = "Green"
                }
                else {
                    ${$menuActivity}.ForeColor = "Red"
                }
            }
            else {
                if ($accTime -gt $requiredTime) {
                    ${$menuActivity}.ForeColor = "Red"
                }
                else {
                    ${$menuActivity}.ForeColor = "Green"
                }
            }
            [void]$contextMenu.Items.Add(${$menuActivity})
            $menuIndexHT[$todayActivity.Activity] = $index
            $index++
            ${$menuActivity}.Add_Click({
                    param($sender, $eventArgs)
                    showActivityTime -menu $sender
                })
        }

    
        $btnTimerActivity = [System.Windows.Forms.Button]::new()
        Add-Member -InputObject $btnTimerActivity -MemberType NoteProperty -Name "Menu" -Value $null
        $btnTimerActivity.Font = $defaultFont
        $btnTimerActivity.AutoSize = $false
        $btnTimerActivity.Text = "No Active Timer"
        Set-ButtonSize $btnTimerActivity
        Set-ActivityTimer -active $false
        $btnTimerActivity.ForeColor = "Pink"
        $btnTimerActivity.Add_Click{
            if ($null -ne $btnTimerActivity.Menu) {
                showActivityTime -menu $btnTimerActivity.Menu
            }
        }
    
        $btnTimerActivityX = 1
        $formPaddingX = 0
        $formPaddingY = 5
        $btnTimerActivity.Location = [System.Drawing.Point]::new($btnTimerActivityX, $formPaddingY)
        $form.Controls.Add($btnTimerActivity)
        $form.PerformLayout()

        $lblSessionTimer = [System.Windows.Forms.Label]::new()
        $lblSessionTimerX = $btnTimerActivity.Size.Width + $btnTimerActivity.Location.X + $formPaddingX
        $lblSessionTimer.AutoSize = $false
        $lblSessionTimer.Font = $defaultFont
        $lblSessionTimer.Text = ""
        $lblSessionTimer.Size = [System.Drawing.Size]::new(0, 0)
        $lblSessionTimer.ForeColor = "White"
        $lblSessionTimer.Location = [System.Drawing.Point]::new($lblSessionTimerX, $formPaddingY)
        $form.Controls.Add($lblSessionTimer)
        $timer = [System.Windows.Forms.Timer]::new()
        $timer.Stop()
        $timer.Interval = 60001
        $form.PerformLayout()
    
        $lblTotalTimer = [System.Windows.Forms.Label]::new()
        #Add-Member -InputObject $lblTotalTimer -MemberType NoteProperty -Name "AccumulatedTime" -Value [TimeSpan]0
        Add-Member -InputObject $lblTotalTimer -MemberType NoteProperty -Name "TargetTime" -Value ([TimeSpan]::Zero)
        Add-Member -InputObject $lblTotalTimer -MemberType NoteProperty -Name "Activity" -Value $null
        Add-Member -InputObject $lblTotalTimer -MemberType NoteProperty -Name "PreviousTime" -Value ([TimeSpan]::Zero)
        Add-Member -InputObject $lblTotalTimer -MemberType NoteProperty -Name "ElapsedTime" -Value ([TimeSpan]::Zero)
        $lblTotalTimer.AutoSize = $false
        $lblTotalTimer.Size = [System.Drawing.Size]::new(0, 0)
        $lblTotalTimer.Text = ""
        $lblTotalTimer.Font = $defaultFont
        $lblTotalTimer.ForeColor = "Red"
        $lblTotalTimerX = $lblSessionTimer.Size.Width + $lblSessionTimer.Location.X + $formPaddingX
        $lblTotalTimer.Location = [System.Drawing.Point]::new($lblTotalTimerX, $formPaddingY)
        $form.Controls.Add($lblTotalTimer)

        $timer.Add_Tick({
                if ($null -ne $lblTotalTimer.Activity) {
                    $activityStartTime = Get-StartTime $lblTotalTimer.Activity
                    if ($activityStartTime -ne $null) {
                        [timespan]$elapsed = [datetime]::now - $activityStartTime
                        $activityAccumulativeTime = Get-AccumulativeTime $lblTotalTimer.Activity
                        $tempTotalTime = $activityAccumulativeTime + $Elapsed 
                        Set-ActivityMenu -activity $lblTotalTimer.Activity -getAccumulativeTime $false -accumulativeTimeValue $tempTotalTime
                        #add elapsed to previous time(for small essential breaks)
                        [timespan]$elapsed += [timespan]$lblTotalTimer.PreviousTime
                        [timespan]$lblTotalTimer.ElapsedTime = [timespan]$elapsed
                        $lblSessionTimer.Text = [string]$elapsed.ToString("hh\:mm")
                        $lblTotalTimer.Text = "($($tempTotalTime.ToString("hh\:mm"))/$($lblTotalTimer.TargetTime.ToString("hh\:mm")))"
                        if ($lblTotalTimer.Activity -ne "Time Waste" -and $lblTotalTimer.Activity -ne "Essential") {
                            if ($tempTotalTime -ge $lblTotalTimer.TargetTime) {
                                $lblTotalTimer.ForeColor = "Green"
                            }
                            else {
                                $lblTotalTimer.ForeColor = "Red"
                            }
                        }
                        else {
                            if ($tempTotalTime -gt $lblTotalTimer.TargetTime) {
                                $lblTotalTimer.ForeColor = "Red"
                            }
                            else {
                                $lblTotalTimer.ForeColor = "Green"
                            }
                        }
                    }
                }
            })

        $btnActivity = New-Object System.Windows.Forms.Button
        $btnActivity.Text = "Activity"
        $btnActivity.Size = New-Object System.Drawing.Size(150, 40)
        $btnActivityX = $lblTotalTimer.Size.Width + $lblTotalTimer.Location.X + $formPaddingX
        $btnActivity.Location = New-Object System.Drawing.Point($btnActivityX, $formPaddingY)
        $btnActivity.ForeColor = "Pink"
        Set-ActivityButton -active $true
    
        $btnActivity.Add_Click({
                $contextMenu.Show($btnActivity, [System.Drawing.Point]::new(0, $btnActivity.Height))  # Show menu just below the button
            })


        $form.Controls.Add($btnActivity)
        $form.PerformLayout()
        $btnPaddingX = 5
        $btnEssential = [System.Windows.Forms.Button]::new()
        $btnEssential.Text = "Break"
        $btnEssential.Size = New-Object System.Drawing.Size(170, 40)
        $btnEssentialLocationX = $btnActivity.Location.X + $btnActivity.Width + $btnPaddingX
        $btnEssential.Location = [System.Drawing.Point]::new($btnEssentialLocationX, $formPaddingY)
        $btnEssential.ForeColor = "LightSeaGreen" 
        Add-Member -InputObject $btnEssential -MemberType NoteProperty -Name "Pressed" -Value $false
        Add-Member -InputObject $btnEssential -MemberType NoteProperty -Name "PreviousActivity" -Value $null
        $btnEssentialStopwatch = [System.Diagnostics.Stopwatch]::new()
        $btnEssentialStopwatch.stop()
        Add-Member -InputObject $btnEssential -MemberType NoteProperty -Name "Stopwatch" -Value $btnEssentialStopwatch
        $btnEssentialTimer = [System.Windows.Forms.Timer]::new()
        $btnEssentialTimer.Stop()
        $btnEssentialTimer.Interval = 1000
        Add-Member -InputObject $btnEssential -MemberType NoteProperty -Name "Timer" -Value $btnEssentialTimer
 

        $btnEssential.add_Paint({
                param($sender, $e)
                $btn = [System.Windows.Forms.Button]$sender
                [System.Windows.Forms.ControlPaint]::DrawBorder(
                    $e.Graphics, 
                    $btn.ClientRectangle,
                    [System.Drawing.Color]::GreenYellow, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::GreenYellow, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::LightSeaGreen, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::LightSeaGreen, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                )
            })
        $btnEssential.Timer.Add_Tick({
                $btnEssential.Text = $btnEssential.Stopwatch.Elapsed.ToString("hh\:mm\:ss")
            })
        $btnEssential.Add_Click({
                Set-EssentialTimer
            })
        $form.Controls.Add($btnEssential)

        $btnChart = [System.Windows.Forms.Button]::new()
        $btnChart.Text = "C"
        $btnChart.Size = New-Object System.Drawing.Size(40, 40)
        $btnChart.ForeColor = "Orchid"
        $btnChartLocationX = $btnEssential.Location.X + $btnEssential.Width + $btnPaddingX
        $btnChart.Location = [System.Drawing.Point]::new($btnChartLocationX, $formPaddingY)
        $btnChart.add_Paint({
                param($sender, $e)
                $btn = [System.Windows.Forms.Button]$sender
                [System.Windows.Forms.ControlPaint]::DrawBorder(
                    $e.Graphics, 
                    $btn.ClientRectangle,
                    [System.Drawing.Color]::MediumPurple, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::MediumPurple, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::Purple, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::Purple, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                )
            })
    
        $btnChart.Add_Click({
                Save-Activities
                Create-Dashboard -accumulativeFilePath $todayAccmulativeFilePath -activityFilePath $todayTimeFilePath -targetActivityFilePath $todayConfigFilePath -htmlFilePath $todayReportFilePath
            })
        $form.Controls.Add($btnChart)

        $btnTimeTarget = [System.Windows.Forms.Button]::new()
        $btnTimeTarget.Text = "T"
        $btnTimeTarget.Size = New-Object System.Drawing.Size(40, 40)
        $btnTimeTarget.ForeColor = "AliceBlue"
        $btnTimeTargetLocationX = $btnChart.Location.X + $btnChart.Width + $btnPaddingX
        $btnTimeTarget.Location = [System.Drawing.Point]::new($btnTimeTargetLocationX, $formPaddingY)
        $btnTimeTarget.add_Paint({
                param($sender, $e)
                $btn = [System.Windows.Forms.Button]$sender
                [System.Windows.Forms.ControlPaint]::DrawBorder(
                    $e.Graphics, 
                    $btn.ClientRectangle,
                    [System.Drawing.Color]::SteelBlue, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::SteelBlue, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::DarkBlue, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid,
                    [System.Drawing.Color]::DarkBlue, 1, [System.Windows.Forms.ButtonBorderStyle]::Solid
                )
            })
        $btnTimeTarget.Add_Click({
                Rename-Item -Path $todayConfigFilePath -NewName $todayConfigBackupFilePath -Force
                $form.Close()        
            })
        $form.Controls.Add($btnTimeTarget )

        $form.Add_Closed{
            Write-Host "Closed Initiated"
            if ($null -ne (Get-CurrentActivity)) {
                Get-CurrentActivity | End-Time 
            }
            End-Time "ApplicationRunTime" 
        }
        $RMBContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
        $exit = $RMBContextMenu.Items.Add("Exit")
        $screenMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Screens")
        $RMBContextMenu.Items.Add($screenMenu)
        $screens = [System.Windows.Forms.Screen]::AllScreens
        for ($i = 0; $i -lt $screens.Length; $i++) {
            $screenItem = New-Object System.Windows.Forms.ToolStripMenuItem("Screen " + $i)
            $screenItem | Add-Member -MemberType NoteProperty -Name "Screen" -Value $i
            [void]$screenMenu.DropDownItems.Add($screenItem)
            $screenItem.add_Click({
                    $global:settingJSON.screenLocation = $this.Screen # Use $currentIndex instead of $i
                    $jsonConfig = $global:settingJSON | ConvertTo-Json
                    Write-Host $global:settingPath 
                    $jsonConfig | Set-Content -Path $global:settingPath -Force
                    $form.close()
                })
        }

        $exit.Add_Click({
                $global:status = $false
                $form.close()
            })
        $form.ContextMenuStrip = $RMBContextMenu

        $form.ShowDialog()
        Write-Host "End of Script"
   
    }
    catch {
        Write-Debug "$($_.Exception.Message)"
    }
}








