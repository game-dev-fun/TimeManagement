
#configPath : Templated schedule
#todayConfigPath: Today's schedule 
#initialize: If loading the application for the first time for the day
function Show-Setting {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$configPath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$todayConfigPath,
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [bool]$initialize = $false,
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$backupConfigPath
    )
    #XAML Input
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms

    $XAMLfilePath = Resolve-Path -LiteralPath "$PSScriptRoot\WPF\ActivitySettingsForm.xaml"
    $inputXAML = Get-Content -Path $XAMLfilePath -Raw
    $inputXAML = $inputXAML -replace 'mc:Ignorable="d"', '' -replace "x:N", "N" -replace '^<Win.*', '<Window'
    [XML]$XAML = $inputXAML
    $reader = New-Object System.Xml.XmlNodeReader $XAML
    try {
        $psForm = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        Write-Debug $_.Exception
        throw
    }
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        try {
            Set-Variable -Name "var_$($_.Name)" -Value $psForm.FindName($_.Name) -ErrorAction Stop
        }
        catch {
            throw
        }
    }
    
    #File Picker
    $var_btnLoadAS.IsEnabled = $false
    $var_btnFilePickerAS.Add_Click({
            $var_btnLoadAS.IsEnabled = $false
            $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $fileDialog.ShowDialog()
            if ($fileDialog.FileName -ne "") {
                $var_txtFileLocationAS.Text = $fileDialog.FileName
                $var_btnLoadAS.IsEnabled = $true
            }
        })


    #Creating Data Table & adding columns
    $dataTable = New-Object System.Data.DataTable
    $dataTable.Columns.AddRange(@("Activity", "Hours"))
    $var_dataGridAS.ItemsSource = $dataTable.DefaultView
    $var_dataGridAS.IsReadOnly = $false
    
    # Updating Column Width
    $var_dataGridAS.add_Loaded({ UpdateColumnWidths })
    $psForm.Add_SizeChanged({ UpdateColumnWidths })
    function UpdateColumnWidths {
        if ($var_dataGridAS.ActualWidth -gt 0) {
            # Calculate and set the column widths as percentages of the total width
            $activityWidth = $var_dataGridAS.ActualWidth * 0.75
            $hoursWidth = $var_dataGridAS.ActualWidth * 0.23
            $var_dataGridAS.Columns[0].Width = $activityWidth
            $var_dataGridAS.Columns[1].Width = $hoursWidth
            $var_dataGridAS.UpdateLayout()
        }
    }

    #Deleting data table Row
    $var_btnDeleteAS.Add_Click({
            $selectedRow = $var_dataGridAS.SelectedItem
            if ($null -ne $selectedRow -and $selectedRow.ToString() -ne "{NewItemPlaceholder}") {
                $dataTable = $var_dataGridAS.ItemsSource.Table
                $dataTable.Rows.Remove($selectedRow.Row)  
                $var_dataGridAS.ItemsSource = $dataTable.DefaultView
            }
            else {
                [System.Windows.MessageBox]::Show("No row selected.", "Delete Error")
            }
        })

    #Cancel Changes
    $var_btnCancelAS.Add_Click({ $psForm.Close() })

    #Clear Rows 
    $var_btnClearAS.Add_Click({
            $dataTable.Rows.Clear()
        })
    
    #Save dialog
    function saveDialogConfigFile {
        param(
            [ref]$saveLocationPath
        )
        $initialDirectory = (Resolve-Path -LiteralPath ".").Path
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Title = "Save Config File As"
        $saveDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        $saveDialog.DefaultExt = "json"
        $saveDialog.InitialDirectory = $initialDirectory   
        $dialogResult = $saveDialog.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $saveLocationPath.Value = $saveDialog.FileName.ToString()
        }
        else {
            $saveLocationPath.Value = $null
        }
        
    }
    

    #serialize the data and saves as JSON objects
    function saveFile {
        param(
            [string]$filePath
        )
        $dataTable
        if (-not [string]::IsNullOrWhiteSpace($filePath)) {
            $essentialExist = $false
            $timeWasteExist = $false
            $saveActivities = [System.Collections.ArrayList]::new()
            foreach ($row in $dataTable.Rows) {
                if (-not [string]::IsNullOrWhiteSpace($row["Activity"])) {
                    # Add the row to the activities array
                    $saveActivities += [PSCustomObject]@{
                        Activity = $row["Activity"]
                        Hours    = $row["Hours"]
                    }
                    # Set flags for "Essential" and "Time Waste"
                    if (-not $essentialExist -and $row["Activity"] -eq "Essential") {
                        $essentialExist = $true
                    }
                    if (-not $timeWasteExist -and $row["Activity"] -eq "Time Waste") {
                        $timeWasteExist = $true
                    }
                }
            }
            $config = [PSCustomObject]@{
                "activities" = $saveActivities
            }
            if (-not $essentialExist) {
                $config.activities += [PSCustomObject]@{
                    Activity = "Essential"
                    Hours    = 0
                }
            }
            if (-not $timeWasteExist) {
                $config.activities += [PSCustomObject]@{
                    Activity = "Time Waste"
                    Hours    = 0
                }
            }
            ConvertTo-Json $config | Out-File -FilePath $filePath -Force
            $var_txtFileLocationAS.Text = $filePath
            return $true
        }
        else {
            return $false
        }
    }
    
    #Saving changes to a different file
    $var_btnSaveAS.Add_Click({
            $saveLocationPath = $var_txtFileLocationAS.Text
            if (-not [string]::IsNullOrWhiteSpace($saveLocationPath)) {
                $saveLocationMsg = [System.Windows.MessageBox]::Show(
                    "Do you want to overwrite $saveLocationPath", 
                    "Confirmation", 
                    [System.Windows.MessageBoxButton]::YesNo, 
                    [System.Windows.MessageBoxImage]::Question 
                )
                if ($saveLocationMsg -eq [System.Windows.MessageBoxResult]::No) {
                    $saveLocationPath = $null
                } 
            }
            
            if ([string]::IsNullOrWhiteSpace($saveLocationPath)) {
                saveDialogConfigFile([ref]$saveLocationPath)
            }
           
            try {
                if (saveFile -filePath $saveLocationPath) {
                    [System.Windows.MessageBox]::Show("Save Successful")
                }
                else {
                    [System.Windows.MessageBox]::Show("Save Cancelled", "File Error")
                }
            }
            catch {
                [System.Windows.MessageBox]::Show($_.Exception.Message, "Error Saving")
            }
        })

    #loading data into form
    function loadData {
        [CmdletBinding()]
        param(
            [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [string]$filePath,
            [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [bool]$initialize = $false
        )
        if ([string]::IsNullOrWhiteSpace($filePath) -and $initialize) {
            $var_txtFileLocationAS.Text = $todayConfigPath
        }
        elseif (Test-Path -Path $filePath) {
            try {
                $fileContent = Get-Content -Path $filePath
                $json = $fileContent | ConvertFrom-Json
                if ($initialize) {
                    $var_txtFileLocationAS.Text = $todayConfigPath
                }
                else {
                    $var_txtFileLocationAS.Text = $filePath
                }
                $activites = $json.activities
                $dataTable.Rows.Clear()
                foreach ($activity in $activites) {
                    $newRow = $dataTable.NewRow()
                    $newRow["Activity"] = $activity.Activity
                    $newRow["Hours"] = $activity.Hours
                    $dataTable.Rows.Add($newRow)
                }
            }
            catch {
                [System.Windows.MessageBox]::Show($_.Exception.Message, "Error Loading")
            }    
        }
        else {
            Write-Host "No config File found"
            [System.Windows.MessageBox]::Show("No template config file provided", "File Error")
        }
   
    }

    #Save the config to the config file
    $var_btnOkAS.Add_Click({
            if (saveFile -filePath $todayConfigPath) {
                [System.Windows.MessageBox]::Show("Today Config Succesfull")
                if(Test-Path $backupConfigPath){
                    Remove-Item -Path $backupConfigPath -Force
                }
            }
            else {
                [System.Windows.MessageBox]::Show("Today Config Failed")
            }
        })
    
    $var_btnLoadAS.Add_Click({ loadData -filePath $var_txtFileLocationAS.Text })
    if ($initialize) {
        if(Test-Path $backupConfigPath){
            loadData -filePath $backupConfigPath -initialize $true
        }else{
            loadData -filePath $configPath -initialize $true
        }
    }
    else {
        loadData -filePath $todayConfigPath
    }
    $psForm.ShowDialog()
}


