#Requires -Version 5.1
# =============================================================================
# queue/QueueGUI.psm1  v2.1
# WPF integration layer for the Task Queue page.
# FIXES:
#   - Selection persistence: uses Tag to restore selection after grid refresh
#   - Param edit dialog: fully functional with Edit-QueueTask
#   - Approved verbs only
#   - Dynamic/responsive layout (DataGrid fills all available space)
#   - ComboBox item colors fixed so text is visible
# ASCII only. PS 5.1 compatible.
# =============================================================================
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Script:QG = @{
    Ctrl         = $null
    PollTimer    = $null
    DragSrcIdx   = $null
    AllTaskDefs  = $null   # reference to $Script:AllTasks from parent UI
}

# ---------------------------------------------------------------------------
# PUBLIC: called from AutoBuild.UI.ps1 after window loads
# ---------------------------------------------------------------------------

function Initialize-QueuePage {
    param(
        [Parameter(Mandatory)][System.Windows.Window]$Window,
        [Parameter(Mandatory)][hashtable]$Ctrl,
        [Parameter(Mandatory)][System.Collections.IEnumerable]$AllTasks,
        [Parameter(Mandatory)][string]$EngineRoot
    )
    $Script:QG.Ctrl        = $Ctrl
    $Script:QG.AllTaskDefs = $AllTasks

    # Init data layer
    Initialize-Queue
    Set-QueueRunnerConfig -EngineRoot $EngineRoot -EventHandlers @{
        OnTaskStarted   = { param($t) Invoke-QueueDispatch { Sync-QueueGrid; Sync-QueueStatus } }
        OnTaskCompleted = { param($t) Invoke-QueueDispatch { Sync-QueueGrid; Sync-QueueStatus } }
        OnTaskFailed    = { param($t,$e) Invoke-QueueDispatch { Sync-QueueGrid; Sync-QueueStatus } }
        OnQueueEmpty    = { Invoke-QueueDispatch { Sync-QueueStatus } }
        OnStateChanged  = { Invoke-QueueDispatch { Sync-QueueStatus } }
    }

    # Populate task combo from current AllTasks
    Sync-QueueCombo

    # Wire all controls
    Connect-QueueControls

    # Start runner in paused state — user must click a Run button
    Start-QueueRunner -UseWpfTimer -PollIntervalMs 800 -AutoAdvance $false
    Suspend-Queue

    # 1.5s auto-refresh when page is visible
    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(1500)
    $t.Add_Tick({
        $p = $Script:QG.Ctrl['pageQueue']
        if ($null -ne $p -and $p.Visibility -eq [System.Windows.Visibility]::Visible) {
            Sync-QueueGrid
            Sync-QueueStatus
        }
    })
    $t.Start()
    $Script:QG.PollTimer = $t

    Sync-QueueGrid
}

# ---------------------------------------------------------------------------
# PUBLIC: called when AllTasks list is refreshed by catalog/execute page
# ---------------------------------------------------------------------------

function Update-QueueTaskList {
    param([System.Collections.IEnumerable]$AllTasks)
    $Script:QG.AllTaskDefs = $AllTasks
    Sync-QueueCombo
}

# ---------------------------------------------------------------------------
# PUBLIC: Add a specific task directly to the queue (called from Catalog/Execute)
# Returns $true if added, $false if user cancelled param dialog
# ---------------------------------------------------------------------------

function Add-TaskToQueueFromUI {
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [hashtable]$DefaultParams = @{}
    )
    if ($null -eq $Script:QG.AllTaskDefs) { return $false }
    $taskDef = @($Script:QG.AllTaskDefs) | Where-Object { $_.Name -eq $TaskName } | Select-Object -First 1
    if ($null -eq $taskDef) { return $false }

    $params = $DefaultParams.Clone()

    # If task has defined params, show dialog
    if ($null -ne $taskDef.Params -and $taskDef.Params.Count -gt 0) {
        $result = Show-QueueParamDialog -TaskName $TaskName `
                                        -ParamDefs  $taskDef.Params `
                                        -CurrentValues $params
        if ($null -eq $result) { return $false }  # user cancelled
        $params = $result
    }

    $t = New-QueueTask -Name $taskDef.Name `
                       -Description $taskDef.Description `
                       -ScriptPath $taskDef.Path `
                       -Parameters $params
    Add-QueueTask -Task $t | Out-Null
    Sync-QueueGrid
    return $true
}

# ---------------------------------------------------------------------------
# INTERNAL: populate add-task ComboBox
# ---------------------------------------------------------------------------

function Sync-QueueCombo {
    $c = $Script:QG.Ctrl['cboQueueAddTask']
    if ($null -eq $c) { return }
    $c.Items.Clear()
    if ($null -eq $Script:QG.AllTaskDefs) { return }

    # Use ItemContainerStyle applied by parent XAML for color.
    # Directly set Foreground/Background on each item for safety.
    $fgBrush = [System.Windows.Media.Brushes]::WhiteSmoke
    $bgBrush = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString('#1E2232')

    foreach ($td in $Script:QG.AllTaskDefs) {
        $item             = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content     = $td.Name
        $item.Tag         = $td
        $item.Foreground  = $fgBrush
        $item.Background  = $bgBrush
        $item.Padding     = [System.Windows.Thickness]::new(6,3,6,3)
        [void]$c.Items.Add($item)
    }
    if ($c.Items.Count -gt 0) { $c.SelectedIndex = 0 }
}

# ---------------------------------------------------------------------------
# INTERNAL: wire all button/menu/drag events
# ---------------------------------------------------------------------------

function Connect-QueueControls {
    $C = $Script:QG.Ctrl

    #--- Add task from combo ---
    $b = $C['btnQueueAdd']
    if ($null -ne $b) {
        $b.Add_Click({
            $c   = $Script:QG.Ctrl['cboQueueAddTask']
            $sel = if ($null -ne $c) { $c.SelectedItem } else { $null }
            if ($null -eq $sel -or $null -eq $sel.Tag) { return }
            $taskDef = $sel.Tag
            $params  = @{}
            if ($null -ne $taskDef.Params -and $taskDef.Params.Count -gt 0) {
                $params = Show-QueueParamDialog -TaskName $taskDef.Name -ParamDefs $taskDef.Params -CurrentValues @{}
                if ($null -eq $params) { return }
            }
            $t = New-QueueTask -Name $taskDef.Name -Description $taskDef.Description `
                               -ScriptPath $taskDef.Path -Parameters $params
            Add-QueueTask -Task $t | Out-Null
            Sync-QueueGrid
        })
    }

    #--- Run All ---
    $b = $C['btnQueueRunAll']
    if ($null -ne $b) {
        $b.Add_Click({
            Start-AllQueueTasks
            Sync-QueueGrid; Sync-QueueStatus
        })
    }

    #--- Run Selected ---
    $b = $C['btnQueueRunSelected']
    if ($null -ne $b) {
        $b.Add_Click({
            $g   = $Script:QG.Ctrl['gridQueue']
            $ids = @($g.SelectedItems | ForEach-Object { $_.TaskId })
            if ($ids.Count -gt 0) {
                Start-SelectedQueueTasks -TaskIds $ids
                Sync-QueueGrid; Sync-QueueStatus
            }
        })
    }

    #--- Next ---
    $b = $C['btnQueueRunNext']
    if ($null -ne $b) {
        $b.Add_Click({
            if ($Script:Runner.IsStopped) { Start-QueueRunner -UseWpfTimer -AutoAdvance $false }
            Start-NextQueueTask
            Sync-QueueGrid; Sync-QueueStatus
        })
    }

    #--- Pause ---
    $b = $C['btnQueuePause']
    if ($null -ne $b) { $b.Add_Click({ Suspend-Queue; Sync-QueueStatus }) }

    #--- Resume ---
    $b = $C['btnQueueResume']
    if ($null -ne $b) {
        $b.Add_Click({
            if ($Script:Runner.IsStopped) { Start-QueueRunner -UseWpfTimer -AutoAdvance $true }
            Resume-Queue
            Sync-QueueGrid; Sync-QueueStatus
        })
    }

    #--- Stop All ---
    $b = $C['btnQueueStop']
    if ($null -ne $b) {
        $b.Add_Click({
            $ans = [System.Windows.MessageBox]::Show(
                'Stop the runner and cancel the active task?',
                'Confirm Stop', 'YesNo', 'Warning')
            if ($ans -eq 'Yes') {
                Stop-QueueRunner
                Sync-QueueGrid; Sync-QueueStatus
            }
        })
    }

    #--- Save queue ---
    $b = $C['btnQueueSave']
    if ($null -ne $b) {
        $b.Add_Click({
            [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Filter   = 'Queue JSON|*.json'
            $dlg.FileName = 'queue_{0}.json' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
            if ($dlg.ShowDialog() -eq 'OK') {
                $ok = Save-TaskQueue -Path $dlg.FileName
                if ($ok) {
                    [void][System.Windows.MessageBox]::Show("Queue saved to:`n$($dlg.FileName)",
                        'AutoBuild Queue','OK','Information')
                }
            }
        })
    }

    #--- Load queue ---
    $b = $C['btnQueueLoad']
    if ($null -ne $b) {
        $b.Add_Click({
            [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Filter = 'Queue JSON|*.json'
            if ($dlg.ShowDialog() -eq 'OK') {
                $n = Import-TaskQueue -Path $dlg.FileName -OnlyPending
                [void][System.Windows.MessageBox]::Show("$n task(s) loaded.",
                    'AutoBuild Queue','OK','Information')
                Sync-QueueGrid
            }
        })
    }

    #--- Clear queue ---
    $b = $C['btnQueueClear']
    if ($null -ne $b) {
        $b.Add_Click({
            $ans = [System.Windows.MessageBox]::Show(
                'Remove all non-running tasks from the queue?',
                'Confirm Clear', 'YesNo', 'Question')
            if ($ans -eq 'Yes') { Clear-TaskQueue; Sync-QueueGrid }
        })
    }

    #--- Context: Edit params ---
    $m = $C['ctxQueueEdit']
    if ($null -ne $m) {
        $m.Add_Click({
            $g   = $Script:QG.Ctrl['gridQueue']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            if ($null -eq $sel) { return }
            $item = Get-QueueTask -TaskId $sel.TaskId
            if ($null -eq $item) { return }
            if ($item.Status -eq 'Running') {
                [void][System.Windows.MessageBox]::Show('Cannot edit a running task.','AutoBuild Queue','OK','Warning')
                return
            }

            # Find param definitions for this task
            $paramDefs = @()
            if ($null -ne $Script:QG.AllTaskDefs) {
                $td = @($Script:QG.AllTaskDefs) | Where-Object { $_.Name -eq $item.Name } | Select-Object -First 1
                if ($null -ne $td -and $null -ne $td.Params) { $paramDefs = $td.Params }
            }

            $newParams = Show-QueueParamDialog `
                -TaskName      $item.Name `
                -ParamDefs     $paramDefs `
                -CurrentValues $item.Parameters

            if ($null -ne $newParams) {
                Edit-QueueTask -TaskId $sel.TaskId -Parameters $newParams | Out-Null
                Sync-QueueGrid
                [void][System.Windows.MessageBox]::Show(
                    "Parameters updated for '$($item.Name)'.",
                    'AutoBuild Queue','OK','Information')
            }
        })
    }

    #--- Context: Run selected ---
    $m = $C['ctxQueueRunSel']
    if ($null -ne $m) {
        $m.Add_Click({
            $g   = $Script:QG.Ctrl['gridQueue']
            $ids = @($g.SelectedItems | ForEach-Object { $_.TaskId })
            if ($ids.Count -gt 0) {
                Start-SelectedQueueTasks -TaskIds $ids
                Sync-QueueGrid; Sync-QueueStatus
            }
        })
    }

    #--- Context: Skip ---
    $m = $C['ctxQueueSkip']
    if ($null -ne $m) {
        $m.Add_Click({
            $g   = $Script:QG.Ctrl['gridQueue']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            if ($null -ne $sel) { Skip-QueueTask -TaskId $sel.TaskId; Sync-QueueGrid }
        })
    }

    #--- Context: Cancel active ---
    $m = $C['ctxQueueCancel']
    if ($null -ne $m) { $m.Add_Click({ Stop-ActiveTask; Sync-QueueGrid; Sync-QueueStatus }) }

    #--- Context: Move Up ---
    $m = $C['ctxQueueMoveUp']
    if ($null -ne $m) {
        $m.Add_Click({
            $g   = $Script:QG.Ctrl['gridQueue']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            if ($null -ne $sel) { Move-QueueTaskUp -TaskId $sel.TaskId; Sync-QueueGrid -SelectTaskId $sel.TaskId }
        })
    }

    #--- Context: Move Down ---
    $m = $C['ctxQueueMoveDown']
    if ($null -ne $m) {
        $m.Add_Click({
            $g   = $Script:QG.Ctrl['gridQueue']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            if ($null -ne $sel) { Move-QueueTaskDown -TaskId $sel.TaskId; Sync-QueueGrid -SelectTaskId $sel.TaskId }
        })
    }

    #--- Context: Move to top ---
    $m = $C['ctxQueueMoveFirst']
    if ($null -ne $m) {
        $m.Add_Click({
            $g   = $Script:QG.Ctrl['gridQueue']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            if ($null -ne $sel) { Move-QueueTaskToPosition -TaskId $sel.TaskId -Position 1; Sync-QueueGrid -SelectTaskId $sel.TaskId }
        })
    }

    #--- Context: Move to bottom ---
    $m = $C['ctxQueueMoveLast']
    if ($null -ne $m) {
        $m.Add_Click({
            $g     = $Script:QG.Ctrl['gridQueue']
            $sel   = if ($null -ne $g) { $g.SelectedItem } else { $null }
            $count = @(Get-QueueSnapshot).Count
            if ($null -ne $sel -and $count -gt 0) {
                Move-QueueTaskToPosition -TaskId $sel.TaskId -Position $count
                Sync-QueueGrid -SelectTaskId $sel.TaskId
            }
        })
    }

    #--- Context: Remove ---
    $m = $C['ctxQueueRemove']
    if ($null -ne $m) {
        $m.Add_Click({
            $g   = $Script:QG.Ctrl['gridQueue']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            if ($null -ne $sel) {
                $ok = Remove-QueueTask -TaskId $sel.TaskId
                if (-not $ok) {
                    [void][System.Windows.MessageBox]::Show(
                        'Cannot remove: task is currently running.',
                        'AutoBuild Queue','OK','Warning')
                }
                Sync-QueueGrid
            }
        })
    }

    #--- Drag and drop reorder ---
    $g = $C['gridQueue']
    if ($null -ne $g) {
        $g.Add_PreviewMouseLeftButtonDown({
            param($sender,$e)
            $dep = $e.OriginalSource -as [System.Windows.DependencyObject]
            while ($null -ne $dep -and $dep -isnot [System.Windows.Controls.DataGridRow]) {
                $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
            }
            $Script:QG.DragSrcIdx = if ($dep -is [System.Windows.Controls.DataGridRow]) { $dep.GetIndex() } else { $null }
        })

        $g.Add_PreviewMouseMove({
            param($sender,$e)
            if ($e.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }
            if ($null -eq $Script:QG.DragSrcIdx) { return }
            [System.Windows.DragDrop]::DoDragDrop(
                $sender,
                "qrow:$($Script:QG.DragSrcIdx)",
                [System.Windows.DragDropEffects]::Move) | Out-Null
        })

        $g.Add_Drop({
            param($sender,$e)
            if (-not $e.Data.GetDataPresent('System.String')) { return }
            $raw = $e.Data.GetData('System.String') -as [string]
            if ($raw -notmatch '^qrow:(\d+)$') { return }
            $srcIdx = [int]$Matches[1]
            $dep    = $e.OriginalSource -as [System.Windows.DependencyObject]
            while ($null -ne $dep -and $dep -isnot [System.Windows.Controls.DataGridRow]) {
                $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
            }
            $dstIdx = if ($dep -is [System.Windows.Controls.DataGridRow]) { $dep.GetIndex() } else { $srcIdx }
            if ($srcIdx -ne $dstIdx -and $srcIdx -ge 0 -and $dstIdx -ge 0) {
                $snap = Get-QueueSnapshot
                if ($srcIdx -lt $snap.Count -and $dstIdx -lt $snap.Count) {
                    $movedId = $snap[$srcIdx].TaskId
                    Move-QueueTaskToPosition -TaskId $movedId -Position ($dstIdx + 1)
                    Sync-QueueGrid -SelectTaskId $movedId
                }
            }
            $Script:QG.DragSrcIdx = $null
        })

        $g.Add_DragOver({
            param($sender,$e)
            $e.Effects = [System.Windows.DragDropEffects]::Move
            $e.Handled = $true
        })
    }
}

# ---------------------------------------------------------------------------
# PUBLIC: Refresh the DataGrid, preserving selection by TaskId
# ---------------------------------------------------------------------------

function Sync-QueueGrid {
    param([string]$SelectTaskId = '')
    $g = $Script:QG.Ctrl['gridQueue']
    if ($null -eq $g) { return }

    # Preserve existing selection if caller didn't specify
    if ([string]::IsNullOrWhiteSpace($SelectTaskId) -and $null -ne $g.SelectedItem) {
        $SelectTaskId = $g.SelectedItem.TaskId
    }

    $rows = @(Get-QueueSnapshot)
    $g.ItemsSource = $rows

    # Restore selection by TaskId
    if (-not [string]::IsNullOrWhiteSpace($SelectTaskId)) {
        for ($i = 0; $i -lt $rows.Count; $i++) {
            if ($rows[$i].TaskId -eq $SelectTaskId) {
                $g.SelectedIndex = $i
                $g.ScrollIntoView($g.SelectedItem)
                break
            }
        }
    }
}

# ---------------------------------------------------------------------------
# PUBLIC: Update the status bar indicators
# ---------------------------------------------------------------------------

function Sync-QueueStatus {
    $state      = Get-QueueRunnerState
    $elStatus   = $Script:QG.Ctrl['elQueueStatus']
    $txtStatus  = $Script:QG.Ctrl['txtQueueStatus']
    $txtActive  = $Script:QG.Ctrl['txtQueueActiveTask']
    $txtElapsed = $Script:QG.Ctrl['txtQueueElapsed']
    $txtOut     = $Script:QG.Ctrl['txtQueueOutput']
    $svOut      = $Script:QG.Ctrl['svQueueOutput']
    $txtTitle   = $Script:QG.Ctrl['txtQueueOutputTitle']

    $green = [System.Windows.Media.Brushes]::LimeGreen
    $gold  = [System.Windows.Media.Brushes]::Goldenrod
    $gray  = [System.Windows.Media.Brushes]::Gray
    $blue  = [System.Windows.Media.Brushes]::DodgerBlue

    if ($state.IsStopped) {
        if ($null -ne $elStatus)  { $elStatus.Background = $gray }
        if ($null -ne $txtStatus) { $txtStatus.Text = 'Stopped' }
    } elseif ($state.IsPaused) {
        if ($null -ne $elStatus)  { $elStatus.Background = $gold }
        if ($null -ne $txtStatus) { $txtStatus.Text = 'Paused' }
    } elseif ($null -ne $state.ActiveTask) {
        if ($null -ne $elStatus)  { $elStatus.Background = $green }
        if ($null -ne $txtStatus) { $txtStatus.Text = 'Running' }
    } else {
        if ($null -ne $elStatus)  { $elStatus.Background = $blue }
        if ($null -ne $txtStatus) { $txtStatus.Text = "Idle ($($state.PendingCount) pending)" }
    }

    if ($null -ne $state.ActiveTask) {
        if ($null -ne $txtActive)  { $txtActive.Text  = "  >  $($state.ActiveTask.Name)" }
        if ($null -ne $txtElapsed) { $txtElapsed.Text = "$($state.ActiveTask.Elapsed)s" }
        if ($null -ne $txtOut)     { $txtOut.Text     = $state.ActiveTask.Output }
        if ($null -ne $svOut)      { $svOut.ScrollToEnd() }
        if ($null -ne $txtTitle)   { $txtTitle.Text   = "| $($state.ActiveTask.Name)" }
    } else {
        if ($null -ne $txtActive)  { $txtActive.Text  = '' }
        if ($null -ne $txtElapsed) { $txtElapsed.Text = '' }
        if ($null -ne $txtTitle)   { $txtTitle.Text   = '' }
    }
}

# ---------------------------------------------------------------------------
# PUBLIC: WPF modal dialog for editing task parameters
# ---------------------------------------------------------------------------

function Show-QueueParamDialog {
    param(
        [string]$TaskName             = '',
        [object[]]$ParamDefs          = @(),
        [hashtable]$CurrentValues     = @{}
    )
    Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Edit Parameters: $TaskName"
        Width="480" SizeToContent="Height" MinHeight="160" MaxHeight="620"
        WindowStartupLocation="CenterOwner"
        Background="#171B26" ResizeMode="CanResizeWithGrip">
  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" MaxHeight="460">
      <StackPanel x:Name="pnlParams" Margin="0,0,0,8"/>
    </ScrollViewer>
    <StackPanel Grid.Row="1" Orientation="Horizontal"
                HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button x:Name="btnOK"     Content="OK"
              Width="90" Height="34" IsDefault="True"
              Background="#F5A623" Foreground="#0F1117"
              FontWeight="Bold" BorderThickness="0" Margin="0,0,10,0"/>
      <Button x:Name="btnCancel" Content="Cancel"
              Width="90" Height="34" IsCancel="True"
              Background="#2D3348" Foreground="#E8EAF0"
              BorderBrush="#3D4460" BorderThickness="1"/>
    </StackPanel>
  </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $win    = [Windows.Markup.XamlReader]::Load($reader)
    $panel  = $win.FindName('pnlParams')
    $btnOK  = $win.FindName('btnOK')
    $btnCan = $win.FindName('btnCancel')

    $darkBg = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString('#1E2232')
    $border = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString('#2D3348')
    $fgPrim = [System.Windows.Media.Brushes]::WhiteSmoke
    $fgSec  = [System.Windows.Media.Brushes]::DarkGray

    $controls = @{}   # paramName -> TextBox

    # Build form from defined params
    foreach ($pd in $ParamDefs) {
        $req = if ($pd.Required) { ' *' } else { '' }
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text       = "$($pd.Name)$req"
        $lbl.Foreground = $fgSec
        $lbl.FontSize   = 11
        $lbl.Margin     = [System.Windows.Thickness]::new(0,10,0,3)
        [void]$panel.Children.Add($lbl)

        $tb = New-Object System.Windows.Controls.TextBox
        $tb.Height      = 32
        $tb.FontSize    = 13
        $tb.Tag         = $pd.Name
        $tb.Background  = $darkBg
        $tb.Foreground  = $fgPrim
        $tb.BorderBrush = $border
        $tb.Padding     = [System.Windows.Thickness]::new(8,4,8,4)
        if ($CurrentValues.ContainsKey($pd.Name)) { $tb.Text = $CurrentValues[$pd.Name] }
        if ($pd.Help)  { $tb.ToolTip = "$($pd.Help)  (Type: $($pd.Type))" }
        [void]$panel.Children.Add($tb)
        $controls[$pd.Name] = $tb

        if ($pd.Help) {
            $h = New-Object System.Windows.Controls.TextBlock
            $h.Text       = $pd.Help
            $h.Foreground = $fgSec
            $h.FontSize   = 10
            $h.Margin     = [System.Windows.Thickness]::new(0,2,0,0)
            [void]$panel.Children.Add($h)
        }
    }

    # Any extra values not in param definitions (free-form)
    foreach ($k in $CurrentValues.Keys) {
        if ($controls.ContainsKey($k)) { continue }
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text       = $k
        $lbl.Foreground = $fgSec
        $lbl.FontSize   = 11
        $lbl.Margin     = [System.Windows.Thickness]::new(0,10,0,3)
        [void]$panel.Children.Add($lbl)

        $tb = New-Object System.Windows.Controls.TextBox
        $tb.Height      = 32
        $tb.FontSize    = 13
        $tb.Tag         = $k
        $tb.Background  = $darkBg
        $tb.Foreground  = $fgPrim
        $tb.BorderBrush = $border
        $tb.Text        = $CurrentValues[$k]
        [void]$panel.Children.Add($tb)
        $controls[$k] = $tb
    }

    # If no params at all, show placeholder
    if ($panel.Children.Count -eq 0) {
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text       = 'This task has no configurable parameters.'
        $lbl.Foreground = $fgSec
        $lbl.Margin     = [System.Windows.Thickness]::new(0,4,0,0)
        [void]$panel.Children.Add($lbl)
    }

    # OK closes with result
    $resultRef = $null
    $btnOK.Add_Click({
        $out = @{}
        foreach ($tb in @($panel.Children | Where-Object { $_ -is [System.Windows.Controls.TextBox] })) {
            if ($null -ne $tb.Tag -and -not [string]::IsNullOrWhiteSpace($tb.Tag)) {
                $out[$tb.Tag.ToString()] = $tb.Text
            }
        }
        $script:resultRef = $out
        $win.DialogResult = $true
        $win.Close()
    })
    $btnCan.Add_Click({ $win.DialogResult = $false; $win.Close() })

    $dlgOK = $win.ShowDialog()
    if ($dlgOK -eq $true) { return $script:resultRef }
    return $null
}

# ---------------------------------------------------------------------------
# INTERNAL: dispatch action on WPF UI thread
# ---------------------------------------------------------------------------

function Invoke-QueueDispatch {
    param([scriptblock]$Action)
    try {
        $g = $Script:QG.Ctrl['gridQueue']
        if ($null -ne $g -and $null -ne $g.Dispatcher) {
            $g.Dispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Normal,
                $Action)
        } else {
            & $Action
        }
    } catch {}
}

Export-ModuleMember -Function @(
    'Initialize-QueuePage',
    'Update-QueueTaskList',
    'Add-TaskToQueueFromUI',
    'Sync-QueueGrid',
    'Sync-QueueStatus',
    'Show-QueueParamDialog',
    'Invoke-QueueDispatch'
)
