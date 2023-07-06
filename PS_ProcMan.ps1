Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

<################################

Poor man's Task Manager via PS GUI

-Get's running process and displays in list view grid.
-Process file locations can be opened by double-clicking the process row.
-Processes can be killed from context menu option (killed via PID(s))

#>

#region HIDE CONSOLE WINDOW
#Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)
#endregion

# Create a form with a tabbed layout
$form = New-Object System.Windows.Forms.Form
$form.Text = "PS Task Manager"
$form.Size = New-Object System.Drawing.Size(575, 700)
$form.StartPosition = "CenterScreen"

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

# Create the "Processes" tab
$processesTabPage = New-Object System.Windows.Forms.TabPage
$processesTabPage.Text = "Processes"

$getButton = New-Object System.Windows.Forms.Button
$getButton.Text = "View"
$getButton.Size = New-Object System.Drawing.Size(50, 30)
$getButton.Location = New-Object System.Drawing.Point(10, 15)
$getButton.Font             = 'Microsoft Sans Serif,12'
$getButton.FlatStyle        = 'Popup'
$getButton.FlatAppearance.BorderSize = 1
$getButton.BackColor        = '#FFAFDAFF'
$getButton.UseVisualStyleBackColor   = $false

$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Location = New-Object System.Drawing.Point(65, 10)
$groupBox1.Size = New-Object System.Drawing.Size(465, 40)
$groupBox1.Text = "Options"

$groupByRadio = New-Object System.Windows.Forms.RadioButton
$groupByRadio.Location = New-Object System.Drawing.Point(10, 15)
$groupByRadio.Size = New-Object System.Drawing.Size(75, 20)
$groupByRadio.Text = "Group By"
$groupByRadio.Checked = $false
$groupBox1.Controls.Add($groupByRadio)


########################
#region LISTVIEW BOX ###
########################
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Location = New-Object System.Drawing.Point(10, 50)
$listView.Size = New-Object System.Drawing.Size(520, 480)
$listView.Font = New-Object System.Drawing.Font("Consolas", 10)
$listView.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
$listView.ShowItemToolTips = $true

$listView.Columns.Add("Name", 160) | Out-Null
$listView.Columns.Add("ID", 50) | Out-Null
$listView.Columns.Add("Mem (MB)", 68) | Out-Null
$listView.Columns.Add("App Title", 200) | Out-Null

# Add event handler for double-click on listview row
$listView.Add_DoubleClick({
    $selectedItem = $listView.SelectedItems[0]
    $processName = $selectedItem.Tag.Name
    $process = Get-Process $processName -ErrorAction SilentlyContinue
    if ($process) {
        $filePath = (Get-Process $processName -FileVersionInfo).FileName | Select-Object -First 1
        if ($filePath -ne '') {
            Start-Process (Split-Path $filePath)
        }
    }
})

####################
####################
#endregion

########################
#region CONTEXT MENU ###
########################
# Create a context menu for sorting
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
# Item 1
$contextMenu.Items.Add("Sort by Name") | Out-Null
#$contextMenu.Items.Add("Sort by ID") | Out-Null
#$contextMenu.Items.Add("Sort by Memory") | Out-Null
# Item 2
$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
# Item 3
$contextMenu.Items.Add("Kill Process") | Out-Null

# Add the context menu to the listview control
$listView.ContextMenuStrip = $contextMenu

# Add event handlers for sorting
$contextMenu.Items[0].Add_Click({ $listView.Sorting = [System.Windows.Forms.SortOrder]::Ascending; $listView.Sort() })
#$contextMenu.Items[1].Add_Click({ $listView.Sorting = [System.Windows.Forms.SortOrder]::Ascending; $listView.ListViewItemSorter = New-Object System.Windows.Forms.ListViewItemComparer(1); $listView.Sort() })
#$contextMenu.Items[2].Add_Click({ $listView.Sorting = [System.Windows.Forms.SortOrder]::Descending; $listView.ListViewItemSorter = New-Object System.Windows.Forms.ListViewItemComparer(2); $listView.Sort() })

########################
########################
#endregion

########################
#region EVENTS
#######################

# CONTEXT MENU EVENTS
#Process that cannot be killed 
$exclusionNames = @("svchost", "system")
$exceptionNames = @("powershell")
$exclusionStrings = @("Operating System","for windows")

$contextMenu.Items[2].Add_Click({ 
    $process = $listview.SelectedItems[0]
    $processId = $process.SubItems[1].Text

    foreach ($excludedWord in $exclusionStrings ) {
        $excludeStringFound = $false
        if (($($process.Tag.Description) -like "*$excludedWord*")  -and ($exceptionNames -notcontains $($process.Tag.Name))) {
            $excludeStringFound = $true
            $denialMessage = "System Related Process"
            break
        }    
    }
    foreach ($excludedName in $exclusionNames ) {
        $excludeNameFound = $false
        if (($($process.Tag.Name) -eq "$excludedName")) {
            $excludeNameFound = $true
            $denialMessage = "Restricted Process"
            break
        }    
    }

    if (($excludeStringFound -eq $false) -and ($excludeNameFound -eq $false)) {
        $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to end process: `r`n`nName: $($process.Tag.Name) `r`nDescription: $($process.Tag.Description) ($($process.Tag.Title))", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            foreach ($ID in $processId.split(",")) {
                Write-Host "stopping............"
                Stop-Process -Id $ID -Force
            }
        } else {
            Write-Host "operation is canceled"
            return
        }

    } else {
        [System.Windows.Forms.MessageBox]::Show("Cannot stop process: `r $($process.Tag.Name) ($($process.Tag.Description)) `r`n `r`n Reason: `r $denialMessage", "Denied", `
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }    
})




# Tooltip
# Add a tooltip to the ListView control
#$toolTip = New-Object System.Windows.Forms.ToolTip
#$tooltip.InitialDelay = 500
#$tooltip.AutoPopDelay = 5000
#$listView.Add_MouseHover({ 
#    $selectedItem = $listView.SelectedItems[0]
#    $processDescript = $selectedItem.Tag.Description
#    Write-Host $processDescript
#    $toolTip.SetToolTip($listView, "$processDescript")
#})

# Filter box/label
$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Location = New-Object System.Drawing.Point(10, 545)
$filterLabel.Size = New-Object System.Drawing.Size(55, 30)
$filterLabel.Text = "Filter:"
$filterLabel.Font = New-Object System.Drawing.Font("Consolas", 10)
$filterLabel.AutoSize = $true

$textBox_Filter = New-Object System.Windows.Forms.TextBox
$textBox_Filter.Location = New-Object System.Drawing.Point(68, 545)
$textBox_Filter.Size = New-Object System.Drawing.Size(150, 20)
################

# Count and mem usage label
$foundLabel = New-Object System.Windows.Forms.Label
$foundLabel.Location = New-Object System.Drawing.Point(220, 545)
$foundLabel.Font = New-Object System.Drawing.Font("Consolas", 12)
$foundLabel.AutoSize = $true
$foundLabel.ForeColor = [System.Drawing.Color]::DarkGreen

$totalMemoryLabel = New-Object System.Windows.Forms.Label
$totalMemoryLabel_LocX = $foundLabel.Right + 25
$totalMemoryLabel.Location = New-Object System.Drawing.Point($totalMemoryLabel_LocX, 545)
$totalMemoryLabel.AutoSize = $true
$totalMemoryLabel.Font = New-Object System.Drawing.Font("Consolas", 12)
$totalMemoryLabel.ForeColor = [System.Drawing.Color]::DarkBlue
################

$getButton.Add_Click({
    # Clear the listview control
    $listView.Items.Clear()
    $totalMemory = 0

    if ($groupByRadio.Checked) {

        # Group processes by name and sum up their memory usage
        $processes = Get-Process | Group-Object -Property ProcessName | 
            Select-Object @{Name="ProcessName"; Expression={$_.Name}}, 
            @{Name="Id"; Expression={$_.Group | Select-Object -ExpandProperty Id}},
            @{Name="Description"; Expression={$_.group | Select-Object -ExpandProperty product}},
            @{Name="Memory"; Expression={[math]::Round(($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)}} | Sort-Object memory -Descending
    } else {
        $processes = Get-Process | 
            Select-Object @{Name="ProcessName"; Expression={$_.Name}}, 
            @{Name="Id"; Expression={$_.Id}},
            @{Name="Description"; Expression={$_.Description}},
            @{Name="App_Title"; Expression={$_.MainWindowTitle}},
            @{Name="Memory"; Expression={[math]::Round(($_.WS / 1MB))}} | Sort-Object memory -Descending
    }

    # Add the processes to the listview control
    # Set intitial count
    $count = 0
    foreach ($process in $processes) {

        $count +=1
        
        # Item 0
        $item = New-Object System.Windows.Forms.ListViewItem($process.ProcessName)
        # Save the file name only to a tag (used later for file location)
        $item.Tag = @{
            ID = $process.ID
            Description = $process.Description
            Name = $process.ProcessName
            Memory = $processes.memory
            Title = $process.App_Title
        }

        # Item 1
        if ($groupByRadio.Checked) {
            $item.SubItems.Add($process.Id -join ',')
            #Separate grouped IDs to be counted
            $IDs = ($item.SubItems[1].Text).split(",")
            #Add number of IDs/processes in parenthesis next to process name
            $item.SubItems[0].Text = "$($process.ProcessName)" + " (" + $($IDs.Count) + ")"
            
        } else {
            $item.SubItems.Add($process.Id)
        }

        # Item 2
        $item.SubItems.Add($process.Memory)

        # Item 3
        if ($process.App_Title -ge 1) {
            $item.SubItems.Add("$($item.Tag.Title)")
        } else {
            $item.SubItems.Add("$($item.Tag.Description)")
        }

        $listView.Items.Add($item) | Out-Null

        $totalMemory += $item.SubItems[2].Text
    }

    $item.ToolTipText = "$($item.Tag.Title) ($($item.Tag.Description))"

    # Update the "found" and "total memory" labels with the number of rows found and total memory respectively
    $foundLabel.Text = "$count processes"
    $totalMemoryLabel.Text = "= $($totalMemory.ToString("N2")) MB"

    $textBox_Filter.Focus()
})

$textBox_Filter.Add_TextChanged({
    # Filter the listview control based on the text in the textbox
    $count = 0
    $totalMemory = 0
    foreach ($item in $listView.Items) {
        if ($item.Text -like "*$($textBox_Filter.Text)*" ) {
            $item.ForeColor = [System.Drawing.Color]::Black
            $count += 1
            $totalMemory += $item.SubItems[2].Text
        } else {
            $item.ForeColor = [System.Drawing.Color]::LightGray
        }
    }

    # Update the "found" and "total memory" labels with the number of rows found and total memory respectively
    $foundLabel.Text = "$count processes"
    $totalMemoryLabel.Text = "= $($totalMemory.ToString("N2")) MB"
})

$listView.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $NextIndex = -1

        if ($listView.SelectedItems.Count -gt 0) {
            $listView.SelectedItems[0].Selected = $false
        }

        for ($i = $script:CurrentIndex + 1; $i -lt $ListView.Items.Count; $i++) {
            if ($ListView.Items[$i].ForeColor -eq [System.Drawing.Color]::Black) {
                $NextIndex = $i
                break;
            }
        }
        
        if ($NextIndex -eq -1) {
            for ($i = 0; $i -le $script:CurrentIndex; $i++) {
                if ($ListView.Items[$i].ForeColor -eq [System.Drawing.Color]::Black) {
                    $NextIndex = $i
                    break;
                }
            }
        }
        
        if ($NextIndex -ne -1) {
            $script:CurrentIndex = $NextIndex
            $ListView.Items[$script:CurrentIndex].Selected = $true
            $listView.Focus()
            $ListView.EnsureVisible($script:CurrentIndex)
            
        }
    }
})

$processesTabPage.Controls.Add($getButton)
$processesTabPage.Controls.Add($groupBox1)
$processesTabPage.Controls.Add($listView)
$processesTabPage.Controls.Add($textBox_Filter)
$processesTabPage.Controls.Add($foundLabel)
$processesTabPage.Controls.Add($filterLabel)
$processesTabPage.Controls.Add($totalMemoryLabel)

# Add the "Processes" tab to the tab control
$tabControl.TabPages.Add($processesTabPage)

# Add the tab, tooltip controls to the form
$form.Controls.Add($tabControl)


# Show the form
$form.ShowDialog() | Out-Null