<#
.SYNOPSIS
    Modern GUI-based Disk Space Analyzer.

.DESCRIPTION
    Launches a modern Windows Forms application to analyze disk usage.
    Features:
    - Dark Mode UI
    - Sortable Grid (Automatic)
    - CSV Export
    - Interactive Deletion (Robust)
    - Resizable Window
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Theme Colors ---
$ColorBackground = [System.Drawing.Color]::FromArgb(30, 30, 30)
$ColorPanel = [System.Drawing.Color]::FromArgb(45, 45, 48)
$ColorText = [System.Drawing.Color]::White
$ColorAccent = [System.Drawing.Color]::FromArgb(0, 122, 204)
$ColorDanger = [System.Drawing.Color]::FromArgb(220, 53, 69)
$ColorSuccess = [System.Drawing.Color]::FromArgb(40, 167, 69)
$ColorGridHeader = [System.Drawing.Color]::FromArgb(60, 60, 60)
$ColorGridRow = [System.Drawing.Color]::FromArgb(45, 45, 48)
$ColorGridAlt = [System.Drawing.Color]::FromArgb(55, 55, 58)

# --- Form Setup ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Disk Space Analyzer"
$Form.Size = New-Object System.Drawing.Size(1200, 700)
$Form.MinimumSize = New-Object System.Drawing.Size(800, 500)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = $ColorBackground
$Form.ForeColor = $ColorText
$Form.FormBorderStyle = "Sizable"

# Title Panel
$TopPanel = New-Object System.Windows.Forms.Panel
$TopPanel.Dock = "Top"
$TopPanel.Height = 120
$TopPanel.BackColor = $ColorPanel
$Form.Controls.Add($TopPanel)

$TitleLabel = New-Object System.Windows.Forms.Label
$TitleLabel.Text = "Disk Space Analyzer"
$TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$TitleLabel.ForeColor = $ColorAccent
$TitleLabel.AutoSize = $true
$TitleLabel.Location = New-Object System.Drawing.Point(20, 20)
$TopPanel.Controls.Add($TitleLabel)

$DescLabel = New-Object System.Windows.Forms.Label
$DescLabel.Text = "Scan folders, analyze usage, and clean up space."
$DescLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$DescLabel.ForeColor = [System.Drawing.Color]::LightGray
$DescLabel.AutoSize = $true
$DescLabel.Location = New-Object System.Drawing.Point(25, 55)
$TopPanel.Controls.Add($DescLabel)

$StatsLabel = New-Object System.Windows.Forms.Label
$StatsLabel.Text = ""
$StatsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$StatsLabel.ForeColor = "White"
$StatsLabel.AutoSize = $true
$StatsLabel.Location = New-Object System.Drawing.Point(25, 75)
$TopPanel.Controls.Add($StatsLabel)

$SelectionLabel = New-Object System.Windows.Forms.Label
$SelectionLabel.Text = ""
$SelectionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$SelectionLabel.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$SelectionLabel.AutoSize = $true
$SelectionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$SelectionLabel.Location = New-Object System.Drawing.Point(550, 75)
$TopPanel.Controls.Add($SelectionLabel)

# Pie Chart Panel for Drive Space (uses Paint event for reliable drawing)
$script:PieChartPanel = New-Object System.Windows.Forms.Panel
$script:PieChartPanel.Size = New-Object System.Drawing.Size(80, 80)
$script:PieChartPanel.Location = New-Object System.Drawing.Point(420, 35)
$script:PieChartPanel.BackColor = $ColorPanel
$script:PieChartPanel.Visible = $false
$TopPanel.Controls.Add($script:PieChartPanel)
$script:PieChartPanel.BringToFront()

# Pie chart data storage
$script:PieUsedGB = 0
$script:PieFreeGB = 0

# Paint event handler for pie chart
$script:PieChartPanel.Add_Paint({
        param($sender, $e)
    
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
        $TotalGB = $script:PieUsedGB + $script:PieFreeGB
        if ($TotalGB -le 0) { return }
    
        # Calculate angles
        $UsedAngle = [float](($script:PieUsedGB / $TotalGB) * 360)
        $FreeAngle = [float](360 - $UsedAngle)
    
        # Colors
        $UsedBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 82, 82))
        $FreeBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(76, 217, 100))
        $BorderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 2)
    
        # Draw area
        $Rect = New-Object System.Drawing.Rectangle(5, 5, 65, 65)
    
        # Draw slices
        if ($UsedAngle -gt 0) {
            $g.FillPie($UsedBrush, $Rect, -90, $UsedAngle)
        }
        if ($FreeAngle -gt 0) {
            $g.FillPie($FreeBrush, $Rect, (-90 + $UsedAngle), $FreeAngle)
        }
    
        # Draw border
        $g.DrawEllipse($BorderPen, $Rect)
    
        # Cleanup
        $UsedBrush.Dispose()
        $FreeBrush.Dispose()
        $BorderPen.Dispose()
    })

# Function to update pie chart
function Update-DriveSpacePieChart {
    param(
        [double]$UsedGB,
        [double]$FreeGB
    )
    
    $script:PieUsedGB = $UsedGB
    $script:PieFreeGB = $FreeGB
    $script:PieChartPanel.Visible = $true
    $script:PieChartPanel.Invalidate()
}

# Controls Panel
$ControlsPanel = New-Object System.Windows.Forms.Panel
$ControlsPanel.Dock = "Top"
$ControlsPanel.Height = 60
$ControlsPanel.Padding = New-Object System.Windows.Forms.Padding(20)
$Form.Controls.Add($ControlsPanel)

# Select Button
$SelectButton = New-Object System.Windows.Forms.Button
$SelectButton.Text = "Select Folder"
$SelectButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$SelectButton.BackColor = $ColorAccent
$SelectButton.ForeColor = "White"
$SelectButton.FlatStyle = "Flat"
$SelectButton.FlatAppearance.BorderSize = 0
$SelectButton.Size = New-Object System.Drawing.Size(140, 35)
$SelectButton.Location = New-Object System.Drawing.Point(20, 10)
$SelectButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$ControlsPanel.Controls.Add($SelectButton)

# Refresh Button
$RefreshButton = New-Object System.Windows.Forms.Button
$RefreshButton.Text = "🔄 Refresh"
$RefreshButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$RefreshButton.BackColor = $ColorPanel
$RefreshButton.ForeColor = "White"
$RefreshButton.FlatStyle = "Flat"
$RefreshButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$RefreshButton.Size = New-Object System.Drawing.Size(100, 35)
$RefreshButton.Location = New-Object System.Drawing.Point(170, 10)
$RefreshButton.Visible = $false
$RefreshButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$ControlsPanel.Controls.Add($RefreshButton)

# Store last scanned path for refresh
$script:LastScannedPath = ""

# Search Box
$SearchBox = New-Object System.Windows.Forms.TextBox
$SearchBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$SearchBox.Size = New-Object System.Drawing.Size(200, 30)
$SearchBox.Location = New-Object System.Drawing.Point(280, 12)
$SearchBox.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$SearchBox.ForeColor = [System.Drawing.Color]::LightGray
$SearchBox.BorderStyle = "FixedSingle"
$SearchBox.Text = "🔍 Search folders..."
$SearchBox.Visible = $false
$ControlsPanel.Controls.Add($SearchBox)

# Search box placeholder behavior
$SearchBox.Add_GotFocus({
        if ($SearchBox.Text -eq "🔍 Search folders...") {
            $SearchBox.Text = ""
            $SearchBox.ForeColor = [System.Drawing.Color]::White
        }
    })

$SearchBox.Add_LostFocus({
        if ($SearchBox.Text -eq "") {
            $SearchBox.Text = "🔍 Search folders..."
            $SearchBox.ForeColor = [System.Drawing.Color]::LightGray
        }
    })

# Live search filtering
$SearchBox.Add_TextChanged({
        $searchText = $SearchBox.Text
        if ($searchText -eq "🔍 Search folders..." -or [string]::IsNullOrWhiteSpace($searchText)) {
            if ($script:GridBindingSource) {
                $script:GridBindingSource.Filter = ""
            }
        }
        else {
            if ($script:GridBindingSource) {
                $script:GridBindingSource.Filter = "Name LIKE '%$($searchText -replace "'", "''")%'"
            }
        }
    })
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(180, 15)
$ProgressBar.Size = New-Object System.Drawing.Size(400, 20)
$ProgressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$ProgressBar.Visible = $false
$ControlsPanel.Controls.Add($ProgressBar)

# Status Label
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready to scan."
$StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$StatusLabel.ForeColor = [System.Drawing.Color]::LightGray
$StatusLabel.AutoSize = $true
$StatusLabel.Location = New-Object System.Drawing.Point(180, 18)
$ControlsPanel.Controls.Add($StatusLabel)

# Select All Button
$SelectAllButton = New-Object System.Windows.Forms.Button
$SelectAllButton.Text = "Select All"
$SelectAllButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$SelectAllButton.BackColor = $ColorPanel
$SelectAllButton.ForeColor = "White"
$SelectAllButton.FlatStyle = "Flat"
$SelectAllButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$SelectAllButton.Size = New-Object System.Drawing.Size(90, 30)
$SelectAllButton.Location = New-Object System.Drawing.Point(710, 12)
$SelectAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$SelectAllButton.Visible = $false
$SelectAllButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$ControlsPanel.Controls.Add($SelectAllButton)

# Unselect All Button
$UnselectAllButton = New-Object System.Windows.Forms.Button
$UnselectAllButton.Text = "Unselect All"
$UnselectAllButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$UnselectAllButton.BackColor = $ColorPanel
$UnselectAllButton.ForeColor = "White"
$UnselectAllButton.FlatStyle = "Flat"
$UnselectAllButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$UnselectAllButton.Size = New-Object System.Drawing.Size(100, 30)
$UnselectAllButton.Location = New-Object System.Drawing.Point(795, 12)
$UnselectAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$UnselectAllButton.Visible = $false
$UnselectAllButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$ControlsPanel.Controls.Add($UnselectAllButton)

# Bottom Panel
$BottomPanel = New-Object System.Windows.Forms.Panel
$BottomPanel.Dock = "Bottom"
$BottomPanel.Height = 70
$BottomPanel.Padding = New-Object System.Windows.Forms.Padding(20)
$Form.Controls.Add($BottomPanel)

# Export Button
$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Text = "Export CSV"
$ExportButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$ExportButton.BackColor = $ColorPanel
$ExportButton.ForeColor = "White"
$ExportButton.FlatStyle = "Flat"
$ExportButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$ExportButton.Size = New-Object System.Drawing.Size(120, 35)
$ExportButton.Location = New-Object System.Drawing.Point(20, 15)
$ExportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$ExportButton.Visible = $false
$ExportButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$BottomPanel.Controls.Add($ExportButton)

# Copy Button
$CopyButton = New-Object System.Windows.Forms.Button
$CopyButton.Text = "Copy To..."
$CopyButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$CopyButton.BackColor = $ColorSuccess
$CopyButton.ForeColor = "White"
$CopyButton.FlatStyle = "Flat"
$CopyButton.FlatAppearance.BorderSize = 0
$CopyButton.Size = New-Object System.Drawing.Size(110, 35)
$CopyButton.Location = New-Object System.Drawing.Point(480, 15)
$CopyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$CopyButton.Visible = $false
$CopyButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$BottomPanel.Controls.Add($CopyButton)

# Rename Button
$RenameButton = New-Object System.Windows.Forms.Button
$RenameButton.Text = "Rename"
$RenameButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$RenameButton.BackColor = $ColorAccent
$RenameButton.ForeColor = "White"
$RenameButton.FlatStyle = "Flat"
$RenameButton.FlatAppearance.BorderSize = 0
$RenameButton.Size = New-Object System.Drawing.Size(100, 35)
$RenameButton.Location = New-Object System.Drawing.Point(600, 15)
$RenameButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$RenameButton.Visible = $false
$RenameButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$BottomPanel.Controls.Add($RenameButton)

# Delete Button
$DeleteButton = New-Object System.Windows.Forms.Button
$DeleteButton.Text = "Delete Selected"
$DeleteButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$DeleteButton.BackColor = $ColorDanger
$DeleteButton.ForeColor = "White"
$DeleteButton.FlatStyle = "Flat"
$DeleteButton.FlatAppearance.BorderSize = 0
$DeleteButton.Size = New-Object System.Drawing.Size(140, 35)
$DeleteButton.Location = New-Object System.Drawing.Point(710, 15)
$DeleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$DeleteButton.Visible = $false
$DeleteButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$BottomPanel.Controls.Add($DeleteButton)

# Letterboxd Import Button
$LetterboxdButton = New-Object System.Windows.Forms.Button
$LetterboxdButton.Text = "📋 Letterboxd"
$LetterboxdButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$LetterboxdButton.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 136)  # Teal
$LetterboxdButton.ForeColor = "White"
$LetterboxdButton.FlatStyle = "Flat"
$LetterboxdButton.FlatAppearance.BorderSize = 0
$LetterboxdButton.Size = New-Object System.Drawing.Size(115, 30)
$LetterboxdButton.Location = New-Object System.Drawing.Point(460, 12)
$LetterboxdButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$LetterboxdButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$ControlsPanel.Controls.Add($LetterboxdButton)

# Top Movies Button
$TopMoviesButton = New-Object System.Windows.Forms.Button
$TopMoviesButton.Text = "🎬 Top Rentals"
$TopMoviesButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$TopMoviesButton.BackColor = [System.Drawing.Color]::FromArgb(156, 39, 176)  # Purple
$TopMoviesButton.ForeColor = "White"
$TopMoviesButton.FlatStyle = "Flat"
$TopMoviesButton.FlatAppearance.BorderSize = 0
$TopMoviesButton.Size = New-Object System.Drawing.Size(115, 30)
$TopMoviesButton.Location = New-Object System.Drawing.Point(580, 12)
$TopMoviesButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$TopMoviesButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$ControlsPanel.Controls.Add($TopMoviesButton)

# Grid Panel (Center)
$GridPanel = New-Object System.Windows.Forms.Panel
$GridPanel.Dock = "Fill"
$GridPanel.Padding = New-Object System.Windows.Forms.Padding(20, 10, 20, 0)
$Form.Controls.Add($GridPanel)
$GridPanel.BringToFront()

# DataGridView
$Grid = New-Object System.Windows.Forms.DataGridView
$Grid.Dock = "Fill"
$Grid.BackgroundColor = $ColorBackground
$Grid.BorderStyle = "None"
$Grid.CellBorderStyle = "SingleHorizontal"
$Grid.ColumnHeadersBorderStyle = "None"
$Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$Grid.EnableHeadersVisualStyles = $false
$Grid.GridColor = $ColorGridAlt
$Grid.RowHeadersVisible = $false
$Grid.AllowUserToAddRows = $false
$Grid.AllowUserToDeleteRows = $false
$Grid.SelectionMode = "FullRowSelect"
$Grid.MultiSelect = $false
$Grid.AutoSizeColumnsMode = "None"
$Grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$Grid.AutoGenerateColumns = $false
$Grid.ReadOnly = $false
$Grid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
$Grid.Visible = $false

# Grid Styles
$HeaderStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
$HeaderStyle.BackColor = $ColorGridHeader
$HeaderStyle.ForeColor = "White"
$HeaderStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$HeaderStyle.Padding = New-Object System.Windows.Forms.Padding(5)
$Grid.ColumnHeadersDefaultCellStyle = $HeaderStyle

$RowStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
$RowStyle.BackColor = $ColorGridRow
$RowStyle.ForeColor = "White"
$RowStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$RowStyle.SelectionBackColor = $ColorAccent
$RowStyle.SelectionForeColor = "White"
$RowStyle.Padding = New-Object System.Windows.Forms.Padding(5)
$Grid.DefaultCellStyle = $RowStyle
$Grid.RowTemplate.Height = 30

$GridPanel.Controls.Add($Grid)

# Columns
$ColCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$ColCheck.HeaderText = "Select"
$ColCheck.DataPropertyName = "Select"
$ColCheck.Width = 60
$ColCheck.AutoSizeMode = "None"
$ColCheck.SortMode = "Automatic"
$ColCheck.ReadOnly = $false
$ColCheck.FalseValue = $false
$ColCheck.TrueValue = $true
$ColCheck.IndeterminateValue = $false
$ColCheck.ThreeState = $false
$ColCheck.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$ColCheck.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(0)
$ColCheck.Name = "Select"
$Grid.Columns.Add($ColCheck) | Out-Null

$ColName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ColName.HeaderText = "Folder Name"
$ColName.DataPropertyName = "Name"
$ColName.Name = "Name"
$ColName.SortMode = "Automatic"
$ColName.ReadOnly = $true
$ColName.Width = 250
$ColName.MinimumWidth = 100
$Grid.Columns.Add($ColName) | Out-Null

$ColMB = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ColMB.HeaderText = "Size (MB)"
$ColMB.DataPropertyName = "SizeMB"
$ColMB.DefaultCellStyle.Format = "N2"
$ColMB.SortMode = "Automatic"
$ColMB.Name = "SizeMB"
$ColMB.ReadOnly = $true
$ColMB.Width = 100
$ColMB.MinimumWidth = 80
$Grid.Columns.Add($ColMB) | Out-Null

$ColGB = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ColGB.HeaderText = "Size (GB)"
$ColGB.DataPropertyName = "SizeGB"
$ColGB.DefaultCellStyle.Format = "N2"
$ColGB.SortMode = "Automatic"
$ColGB.Name = "SizeGB"
$ColGB.ReadOnly = $true
$ColGB.Width = 100
$ColGB.MinimumWidth = 80
$Grid.Columns.Add($ColGB) | Out-Null

$ColFiles = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ColFiles.HeaderText = "File Count"
$ColFiles.DataPropertyName = "Files"
$ColFiles.SortMode = "Automatic"
$ColFiles.Name = "Files"
$ColFiles.ReadOnly = $true
$ColFiles.Width = 90
$ColFiles.MinimumWidth = 70
$Grid.Columns.Add($ColFiles) | Out-Null

$ColPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ColPath.HeaderText = "Full Path"
$ColPath.DataPropertyName = "Path"
$ColPath.SortMode = "Automatic"
$ColPath.Name = "Path"
$ColPath.ReadOnly = $true
$ColPath.Width = 350
$ColPath.MinimumWidth = 150
$Grid.Columns.Add($ColPath) | Out-Null

$ColModified = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ColModified.HeaderText = "Date Modified"
$ColModified.DataPropertyName = "Modified"
$ColModified.SortMode = "Automatic"
$ColModified.Name = "Modified"
$ColModified.ReadOnly = $true
$ColModified.DefaultCellStyle.Format = "g"
$ColModified.Width = 150
$ColModified.MinimumWidth = 100
$Grid.Columns.Add($ColModified) | Out-Null

# --- Logic ---

# Global DataTable
$Global:DataTable = New-Object System.Data.DataTable
$Global:DataTable.Columns.Add("Select", [bool]) | Out-Null
$Global:DataTable.Columns["Select"].DefaultValue = $false
$Global:DataTable.Columns.Add("Name", [string]) | Out-Null
$Global:DataTable.Columns.Add("SizeMB", [double]) | Out-Null
$Global:DataTable.Columns.Add("SizeGB", [double]) | Out-Null
$Global:DataTable.Columns.Add("Files", [int]) | Out-Null
$Global:DataTable.Columns.Add("Path", [string]) | Out-Null
$Global:DataTable.Columns.Add("Modified", [datetime]) | Out-Null

# Function to update selection summary
function Update-SelectionSummary {
    # Count selected items by iterating through Grid rows (more reliable than DataTable query)
    $TotalFolders = 0
    $TotalFiles = 0
    $TotalSizeGB = 0.0
    
    foreach ($Row in $Grid.Rows) {
        if ($Row.Cells["Select"].Value -eq $true) {
            $TotalFolders++
            $TotalFiles += [int]$Row.Cells["Files"].Value
            $TotalSizeGB += [double]$Row.Cells["SizeGB"].Value
        }
    }
    
    if ($TotalFolders -gt 0) {
        $TotalSizeGB = [math]::Round($TotalSizeGB, 2)
        $SelectionLabel.Text = "Selected: $TotalFolders Folders | $TotalFiles Files | $TotalSizeGB GB"
    }
    else {
        $SelectionLabel.Text = ""
    }
}

# Function to get selected rows from Grid
function Get-SelectedRows {
    $Selected = [System.Collections.ArrayList]::new()
    foreach ($Row in $Grid.Rows) {
        if ($Row.Cells["Select"].Value -eq $true) {
            [void]$Selected.Add([PSCustomObject]@{
                    Name     = $Row.Cells["Name"].Value
                    Path     = $Row.Cells["Path"].Value
                    Files    = [int]$Row.Cells["Files"].Value
                    SizeGB   = [double]$Row.Cells["SizeGB"].Value
                    SizeMB   = [double]$Row.Cells["SizeMB"].Value
                    RowIndex = $Row.Index
                })
        }
    }
    return $Selected
}

# Function to extract movie name from folder name
function Extract-MovieName {
    param([string]$FolderName)
    
    $Name = $FolderName
    
    # Common patterns to remove (order matters - more specific first)
    $PatternsToRemove = @(
        # Resolution patterns
        '\b(2160p|1080p|720p|480p|4K|UHD|HD|SD)\b',
        # Video codecs
        '\b(x264|x265|h\.?264|h\.?265|HEVC|AVC|XviD|DivX)\b',
        # Audio codecs
        '\b(AAC|AC3|DTS|DTS-HD|TrueHD|FLAC|MP3|DD5\.?1|5\.1|7\.1|Atmos)\b',
        # Source types
        '\b(BluRay|Blu-Ray|BDRip|BRRip|HDRip|DVDRip|WEBRip|WEB-DL|WEBDL|WEB|HDTV|DVDScr|CAM|TS|R5|DVDRip)\b',
        # Release groups (in brackets or after dash)
        '\[.*?\]',
        '\{.*?\}',
        '[-\s]+[A-Z0-9]{2,}$',
        # Common tags
        '\b(EXTENDED|UNRATED|REMASTERED|DIRECTORS\.?CUT|THEATRICAL|IMAX|3D|HDR|HDR10|DV|Proper|REPACK)\b',
        # File extensions that might be in folder names
        '\.(mkv|mp4|avi|mov|wmv)$',
        # Quality indicators
        '\b(HQ|LQ|REMUX)\b',
        # Subtitles indicators
        '\b(SUBBED|DUBBED|MULTI|ENG|SPA|FRE|GER|ITA|POR|RUS|JPN|KOR|CHI)\b'
    )
    
    # Remove year pattern but keep track of it
    $YearMatch = [regex]::Match($Name, '\(?(19|20)\d{2}\)?')
    $Year = ""
    if ($YearMatch.Success) {
        $Year = $YearMatch.Value -replace '[()]', ''
        # Remove year and everything after it (often the technical stuff follows the year)
        $YearIndex = $Name.IndexOf($YearMatch.Value)
        if ($YearIndex -gt 5) {
            $Name = $Name.Substring(0, $YearIndex)
        }
    }
    
    # Apply removal patterns
    foreach ($Pattern in $PatternsToRemove) {
        $Name = $Name -replace "(?i)$Pattern", ''
    }
    
    # Replace common separators with spaces
    $Name = $Name -replace '[._]', ' '
    
    # Remove extra spaces and trim
    $Name = ($Name -replace '\s+', ' ').Trim()
    
    # Remove trailing dashes or dots
    $Name = $Name -replace '[-.\s]+$', ''
    
    # Title case the result
    $TextInfo = (Get-Culture).TextInfo
    $Name = $TextInfo.ToTitleCase($Name.ToLower())
    
    # Add year back if found
    if ($Year) {
        $Name = "$Name ($Year)"
    }
    
    return $Name
}

# Function to remove a movie from Radarr monitoring when folder is deleted
function Remove-MovieFromRadarr {
    param(
        [string]$FolderPath,
        [string]$FolderName
    )
    
    $ConfigPath = Join-Path $env:APPDATA "RadarrConfig.json"
    if (-not (Test-Path $ConfigPath)) { return $null }
    
    try {
        $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        if (-not $Config.ApiKey -or -not $Config.RadarrUrl) { return $null }
        
        $RadarrUrl = $Config.RadarrUrl.TrimEnd('/')
        $headers = @{ "X-Api-Key" = $Config.ApiKey }
        
        # Get all movies from Radarr
        $movies = Invoke-RestMethod -Uri "$RadarrUrl/api/v3/movie" -Headers $headers -Method Get -TimeoutSec 10
        
        # Find movie by matching path or folder name
        $movie = $movies | Where-Object { 
            $_.path -eq $FolderPath -or 
            $_.folderName -eq $FolderName -or
            $_.path -like "*\$FolderName" -or 
            $_.path -like "*/$FolderName"
        } | Select-Object -First 1
        
        if ($movie) {
            # Delete from Radarr with import exclusion to prevent re-adding
            $deleteUrl = "$RadarrUrl/api/v3/movie/$($movie.id)?addImportExclusion=true"
            Invoke-RestMethod -Uri $deleteUrl -Headers $headers -Method Delete -TimeoutSec 10 | Out-Null
            return $movie.title
        }
    }
    catch { }
    return $null
}

# Function to show rename dialog
function Show-RenameDialog {
    param(
        [string]$CurrentName,
        [string]$SuggestedName
    )
    
    $DialogForm = New-Object System.Windows.Forms.Form
    $DialogForm.Text = "Rename Folder"
    $DialogForm.Size = New-Object System.Drawing.Size(500, 200)
    $DialogForm.StartPosition = "CenterParent"
    $DialogForm.FormBorderStyle = "FixedDialog"
    $DialogForm.MaximizeBox = $false
    $DialogForm.MinimizeBox = $false
    $DialogForm.BackColor = $ColorBackground
    $DialogForm.ForeColor = "White"
    
    $CurrentLabel = New-Object System.Windows.Forms.Label
    $CurrentLabel.Text = "Current: $CurrentName"
    $CurrentLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $CurrentLabel.ForeColor = [System.Drawing.Color]::LightGray
    $CurrentLabel.Location = New-Object System.Drawing.Point(20, 15)
    $CurrentLabel.Size = New-Object System.Drawing.Size(440, 20)
    $DialogForm.Controls.Add($CurrentLabel)
    
    $NewNameLabel = New-Object System.Windows.Forms.Label
    $NewNameLabel.Text = "New Name:"
    $NewNameLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $NewNameLabel.Location = New-Object System.Drawing.Point(20, 45)
    $NewNameLabel.AutoSize = $true
    $DialogForm.Controls.Add($NewNameLabel)
    
    $NameTextBox = New-Object System.Windows.Forms.TextBox
    $NameTextBox.Text = $SuggestedName
    $NameTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $NameTextBox.Location = New-Object System.Drawing.Point(20, 70)
    $NameTextBox.Size = New-Object System.Drawing.Size(440, 30)
    $NameTextBox.BackColor = $ColorPanel
    $NameTextBox.ForeColor = "White"
    $NameTextBox.BorderStyle = "FixedSingle"
    $DialogForm.Controls.Add($NameTextBox)
    
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "Rename"
    $OKButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $OKButton.BackColor = $ColorSuccess
    $OKButton.ForeColor = "White"
    $OKButton.FlatStyle = "Flat"
    $OKButton.FlatAppearance.BorderSize = 0
    $OKButton.Size = New-Object System.Drawing.Size(100, 35)
    $OKButton.Location = New-Object System.Drawing.Point(250, 115)
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $DialogForm.Controls.Add($OKButton)
    
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = "Cancel"
    $CancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $CancelButton.BackColor = $ColorPanel
    $CancelButton.ForeColor = "White"
    $CancelButton.FlatStyle = "Flat"
    $CancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $CancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $CancelButton.Location = New-Object System.Drawing.Point(360, 115)
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $DialogForm.Controls.Add($CancelButton)
    
    $DialogForm.AcceptButton = $OKButton
    $DialogForm.CancelButton = $CancelButton
    
    $Result = $DialogForm.ShowDialog()
    
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $NameTextBox.Text.Trim()
    }
    return $null
}

# Handle Checkbox Clicks Immediately (Dirty State)
$Grid.Add_CurrentCellDirtyStateChanged({
        if ($Grid.IsCurrentCellDirty) {
            [void]$Grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
            Update-SelectionSummary
        }
    })

# Flag to prevent auto-check during loading
$Global:IsLoading = $false

# Toggle Checkbox on Row Click (handles re-clicking already selected rows)
$Grid.Add_CellClick({
        param($sender, $e)
        if (-not $Global:IsLoading -and $e.RowIndex -ge 0) {
            # Don't toggle if clicking directly on the checkbox column (it handles itself)
            if ($e.ColumnIndex -eq 0) {
                return
            }
        
            $Row = $Grid.Rows[$e.RowIndex]
            # Toggle the checkbox value
            if ($Row.Cells["Select"].Value -eq $true) {
                $Row.Cells["Select"].Value = $false
            }
            else {
                $Row.Cells["Select"].Value = $true
            }
            Update-SelectionSummary
        }
    })

# Context Menu for Right-Click
$ContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenu.BackColor = $ColorPanel
$ContextMenu.ForeColor = "White"

$MenuItemOpen = $ContextMenu.Items.Add("Open Folder Location")
$MenuItemOpen.Add_Click({
        if ($Grid.SelectedRows.Count -gt 0) {
            $Path = $Grid.SelectedRows[0].Cells["Path"].Value
            if ($Path -and (Test-Path $Path)) {
                Invoke-Item $Path
            }
        }
    })

$ContextMenu.Items.Add("-") | Out-Null  # Separator

$MenuItemCopyPath = $ContextMenu.Items.Add("Copy Path to Clipboard")
$MenuItemCopyPath.Add_Click({
        if ($Grid.SelectedRows.Count -gt 0) {
            $Path = $Grid.SelectedRows[0].Cells["Path"].Value
            if ($Path) {
                [System.Windows.Forms.Clipboard]::SetText($Path)
                $StatusLabel.Text = "Path copied to clipboard"
                $StatusLabel.Visible = $true
            }
        }
    })

$ContextMenu.Items.Add("-") | Out-Null  # Separator

$MenuItemRename = $ContextMenu.Items.Add("Rename...")
$MenuItemRename.Add_Click({
        if ($Grid.SelectedRows.Count -gt 0) {
            $RowIndex = $Grid.SelectedRows[0].Index
            $CurrentName = $Grid.SelectedRows[0].Cells["Name"].Value
            $CurrentPath = $Grid.SelectedRows[0].Cells["Path"].Value
        
            $SuggestedName = Extract-MovieName -FolderName $CurrentName
            $NewName = Show-RenameDialog -CurrentName $CurrentName -SuggestedName $SuggestedName
        
            if ($NewName -and $NewName -ne $CurrentName) {
                $InvalidChars = [System.IO.Path]::GetInvalidFileNameChars()
                $HasInvalid = $false
                foreach ($Char in $InvalidChars) {
                    if ($NewName.Contains($Char)) { $HasInvalid = $true; break }
                }
            
                if ($HasInvalid) {
                    [System.Windows.Forms.MessageBox]::Show("Invalid characters in name.", "Error", "OK", "Error") | Out-Null
                    return
                }
            
                $ParentPath = Split-Path -Path $CurrentPath -Parent
                $NewPath = Join-Path -Path $ParentPath -ChildPath $NewName
            
                if (Test-Path $NewPath) {
                    [System.Windows.Forms.MessageBox]::Show("A folder with that name already exists.", "Error", "OK", "Error") | Out-Null
                    return
                }
            
                try {
                    Rename-Item -Path $CurrentPath -NewName $NewName -ErrorAction Stop
                    $Grid.Rows[$RowIndex].Cells["Name"].Value = $NewName
                    $Grid.Rows[$RowIndex].Cells["Path"].Value = $NewPath
                    $Global:DataTable.Rows[$RowIndex]["Name"] = $NewName
                    $Global:DataTable.Rows[$RowIndex]["Path"] = $NewPath
                    $Grid.Refresh()
                    [System.Windows.Forms.MessageBox]::Show("Folder renamed!", "Success", "OK", "Information") | Out-Null
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Failed: $_", "Error", "OK", "Error") | Out-Null
                }
            }
        }
    })

$MenuItemCopy = $ContextMenu.Items.Add("Copy To...")
$MenuItemCopy.Add_Click({
        if ($Grid.SelectedRows.Count -gt 0) {
            $SourcePath = $Grid.SelectedRows[0].Cells["Path"].Value
            $FolderName = $Grid.SelectedRows[0].Cells["Name"].Value
        
            $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $FolderBrowser.Description = "Select Destination Folder"
            $FolderBrowser.ShowNewFolderButton = $true
        
            if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $DestRoot = $FolderBrowser.SelectedPath
                $TargetPath = Join-Path -Path $DestRoot -ChildPath $FolderName
            
                try {
                    Copy-Item -Path $SourcePath -Destination $TargetPath -Recurse -Force -ErrorAction Stop
                    [System.Windows.Forms.MessageBox]::Show("Folder copied successfully!", "Success", "OK", "Information") | Out-Null
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Copy failed: $_", "Error", "OK", "Error") | Out-Null
                }
            }
        }
    })

$ContextMenu.Items.Add("-") | Out-Null  # Separator

$MenuItemDelete = $ContextMenu.Items.Add("Delete")
$MenuItemDelete.Add_Click({
        if ($Grid.SelectedRows.Count -gt 0) {
            $FolderName = $Grid.SelectedRows[0].Cells["Name"].Value
            $FolderPath = $Grid.SelectedRows[0].Cells["Path"].Value
        
            $Confirm = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to PERMANENTLY DELETE:`n`n$FolderName", 
                "Confirm Deletion", "YesNo", "Warning"
            )
        
            if ($Confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                try {
                    Remove-Item -Path $FolderPath -Recurse -Force -ErrorAction Stop
                    
                    # Try to remove from Radarr monitoring
                    $RemovedTitle = Remove-MovieFromRadarr -FolderPath $FolderPath -FolderName $FolderName
                    if ($RemovedTitle) {
                        [System.Windows.Forms.MessageBox]::Show("Folder deleted and '$RemovedTitle' removed from Radarr.", "Success", "OK", "Information") | Out-Null
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show("Folder deleted.", "Success", "OK", "Information") | Out-Null
                    }
                    
                    # Refresh the list
                    $RefreshButton.PerformClick()
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Delete failed: $_", "Error", "OK", "Error") | Out-Null
                }
            }
        }
    })

$Grid.ContextMenuStrip = $ContextMenu

# Handle Right-Click Selection
$Grid.Add_CellMouseDown({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
            $Global:IsLoading = $true
            $Grid.ClearSelection()
            $Grid.Rows[$e.RowIndex].Selected = $true
            $Global:IsLoading = $false
        }
    })

$SelectButton.Add_Click({
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select the folder to scan"
        $FolderBrowser.ShowNewFolderButton = $false

        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $RootPath = $FolderBrowser.SelectedPath
            $script:LastScannedPath = $RootPath

            # UI Prep
            $Global:IsLoading = $true
            $SelectButton.Enabled = $false
            $Global:DataTable.Clear()
            $Grid.DataSource = $null
            $Grid.Visible = $false

            $DeleteButton.Visible = $false
            $ExportButton.Visible = $false
            $CopyButton.Visible = $false
            $SelectAllButton.Visible = $false
            $UnselectAllButton.Visible = $false
            $ProgressBar.Visible = $true
            $StatusLabel.Visible = $true
            $ProgressBar.Value = 0
            $StatusLabel.Text = "Scanning directory structure..."
            $StatsLabel.Text = ""
            $SelectionLabel.Text = ""
            $Form.Refresh()

            # Get Directories
            $SubDirs = Get-ChildItem -Path $RootPath -Directory

            if (-not $SubDirs) {
                [System.Windows.Forms.MessageBox]::Show("No subdirectories found in $RootPath", "Info", "OK", "Information") | Out-Null
                $SelectButton.Enabled = $true
                $ProgressBar.Visible = $false
                $StatusLabel.Text = "Ready"
                $Global:IsLoading = $false
                return
            }

            $TotalDirs = $SubDirs.Count
            $Current = 0
            $TotalSize = 0
            $TotalFiles = 0

            # Scan Loop
            foreach ($Dir in $SubDirs) {
                $Current++
                $Percent = [int](($Current / $TotalDirs) * 100)
                $ProgressBar.Value = $Percent
                $StatusLabel.Text = "Scanning: $($Dir.Name)"

                [System.Windows.Forms.Application]::DoEvents()

                try {
                    $Measure = Get-ChildItem -Path $Dir.FullName -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                    $SizeMB = [math]::Round(($Measure.Sum / 1MB), 2)
                    $SizeGB = [math]::Round(($Measure.Sum / 1GB), 2)

                    $Row = $Global:DataTable.NewRow()
                    $Row["Select"] = $false
                    $Row["Name"] = $Dir.Name
                    $Row["SizeMB"] = $SizeMB
                    $Row["SizeGB"] = $SizeGB
                    $Row["Files"] = $Measure.Count
                    $Row["Path"] = $Dir.FullName
                    $Row["Modified"] = $Dir.LastWriteTime
                    $Global:DataTable.Rows.Add($Row)

                    $TotalSize += $Measure.Sum
                    $TotalFiles += $Measure.Count
                }
                catch { }
            }

            # Bind Data
            $script:GridBindingSource = New-Object System.Windows.Forms.BindingSource
            $script:GridBindingSource.DataSource = $Global:DataTable
            $Grid.DataSource = $script:GridBindingSource
            $Grid.Sort($Grid.Columns[2], [System.ComponentModel.ListSortDirection]::Descending)

            $StatusLabel.Visible = $false
            $ProgressBar.Visible = $false
            $Grid.Visible = $true
            $CopyButton.Visible = $true
            $RenameButton.Visible = $true
            $DeleteButton.Visible = $true
            $ExportButton.Visible = $true
            $SelectAllButton.Visible = $true
            $UnselectAllButton.Visible = $true
            $RefreshButton.Visible = $true
            $SearchBox.Visible = $true
            $SelectButton.Enabled = $true

            # Clear selection AFTER grid is visible to prevent auto-select
            $Grid.ClearSelection()
            $Grid.CurrentCell = $null
            $Global:IsLoading = $false

            $TotalSizeGB = [math]::Round(($TotalSize / 1GB), 2)
            
            # Get drive space info
            $DriveLetter = $RootPath.Substring(0, 1)
            $Drive = Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue
            if ($Drive) {
                $UsedGB = [math]::Round($Drive.Used / 1GB, 1)
                $FreeGB = [math]::Round($Drive.Free / 1GB, 1)
                $TotalDriveGB = [math]::Round(($Drive.Used + $Drive.Free) / 1GB, 1)
                $StatsLabel.Text = "Total: $TotalDirs Folders | $TotalFiles Files | $TotalSizeGB GB scanned`nDrive $($DriveLetter): $UsedGB GB used / $FreeGB GB free ($TotalDriveGB GB total)"
                Update-DriveSpacePieChart -UsedGB $UsedGB -FreeGB $FreeGB
            }
            else {
                $StatsLabel.Text = "Total: $TotalDirs Folders | $TotalFiles Files | $TotalSizeGB GB"
                $script:PieChartPanel.Visible = $false
            }
        }
    })

# Refresh Button Click - Rescan the last folder
$RefreshButton.Add_Click({
        if (-not $script:LastScannedPath -or -not (Test-Path $script:LastScannedPath)) {
            [System.Windows.Forms.MessageBox]::Show("No folder to refresh. Please select a folder first.", "Info", "OK", "Information") | Out-Null
            return
        }
        
        $RootPath = $script:LastScannedPath
        
        # UI Prep
        $Global:IsLoading = $true
        $SelectButton.Enabled = $false
        $RefreshButton.Enabled = $false
        $Global:DataTable.Clear()
        $Grid.DataSource = $null
        $Grid.Visible = $false
        
        $DeleteButton.Visible = $false
        $ExportButton.Visible = $false
        $CopyButton.Visible = $false
        $SelectAllButton.Visible = $false
        $UnselectAllButton.Visible = $false
        $ProgressBar.Visible = $true
        $StatusLabel.Visible = $true
        $ProgressBar.Value = 0
        $StatusLabel.Text = "Refreshing..."
        $StatsLabel.Text = ""
        $SelectionLabel.Text = ""
        $Form.Refresh()
        
        # Get Directories
        $SubDirs = Get-ChildItem -Path $RootPath -Directory
        
        if (-not $SubDirs) {
            [System.Windows.Forms.MessageBox]::Show("No subdirectories found in $RootPath", "Info", "OK", "Information") | Out-Null
            $SelectButton.Enabled = $true
            $RefreshButton.Enabled = $true
            $ProgressBar.Visible = $false
            $StatusLabel.Text = "Ready"
            $Global:IsLoading = $false
            return
        }
        
        $TotalDirs = $SubDirs.Count
        $Current = 0
        $TotalSize = 0
        $TotalFiles = 0
        
        foreach ($Dir in $SubDirs) {
            $Current++
            $Percent = [int](($Current / $TotalDirs) * 100)
            $ProgressBar.Value = $Percent
            $StatusLabel.Text = "Scanning: $($Dir.Name)"
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                $Measure = Get-ChildItem -Path $Dir.FullName -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                $SizeMB = [math]::Round(($Measure.Sum / 1MB), 2)
                $SizeGB = [math]::Round(($Measure.Sum / 1GB), 2)
                
                $Row = $Global:DataTable.NewRow()
                $Row["Select"] = $false
                $Row["Name"] = $Dir.Name
                $Row["SizeMB"] = $SizeMB
                $Row["SizeGB"] = $SizeGB
                $Row["Files"] = $Measure.Count
                $Row["Path"] = $Dir.FullName
                $Row["Modified"] = $Dir.LastWriteTime
                $Global:DataTable.Rows.Add($Row)
                
                $TotalSize += $Measure.Sum
                $TotalFiles += $Measure.Count
            }
            catch { }
        }
        
        # Bind Data
        $script:GridBindingSource = New-Object System.Windows.Forms.BindingSource
        $script:GridBindingSource.DataSource = $Global:DataTable
        $Grid.DataSource = $script:GridBindingSource
        $Grid.Sort($Grid.Columns[2], [System.ComponentModel.ListSortDirection]::Descending)
        
        $StatusLabel.Visible = $false
        $ProgressBar.Visible = $false
        $Grid.Visible = $true
        $CopyButton.Visible = $true
        $RenameButton.Visible = $true
        $DeleteButton.Visible = $true
        $ExportButton.Visible = $true
        $SelectAllButton.Visible = $true
        $UnselectAllButton.Visible = $true
        $RefreshButton.Visible = $true
        $SearchBox.Visible = $true
        $SelectButton.Enabled = $true
        $RefreshButton.Enabled = $true
        
        $Grid.ClearSelection()
        $Grid.CurrentCell = $null
        $Global:IsLoading = $false
        
        $TotalSizeGB = [math]::Round(($TotalSize / 1GB), 2)
        
        $DriveLetter = $RootPath.Substring(0, 1)
        $Drive = Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue
        if ($Drive) {
            $UsedGB = [math]::Round($Drive.Used / 1GB, 1)
            $FreeGB = [math]::Round($Drive.Free / 1GB, 1)
            $TotalDriveGB = [math]::Round(($Drive.Used + $Drive.Free) / 1GB, 1)
            $StatsLabel.Text = "Total: $TotalDirs Folders | $TotalFiles Files | $TotalSizeGB GB scanned`nDrive $($DriveLetter): $UsedGB GB used / $FreeGB GB free ($TotalDriveGB GB total)"
            Update-DriveSpacePieChart -UsedGB $UsedGB -FreeGB $FreeGB
        }
        else {
            $StatsLabel.Text = "Total: $TotalDirs Folders | $TotalFiles Files | $TotalSizeGB GB"
            $script:PieChartPanel.Visible = $false
        }
    })

# Top Movies Button Click - Fetch trending movie rentals with cover art
$TopMoviesButton.Add_Click({
        # Create the popup form
        $MoviesForm = New-Object System.Windows.Forms.Form
        $MoviesForm.Text = "Top 50 Movie Rentals"
        $MoviesForm.Size = New-Object System.Drawing.Size(650, 650)
        $MoviesForm.StartPosition = "CenterScreen"
        $MoviesForm.FormBorderStyle = "FixedDialog"
        $MoviesForm.MaximizeBox = $false
        $MoviesForm.BackColor = $ColorBackground
        $MoviesForm.ForeColor = "White"
    
        $TitleLabel = New-Object System.Windows.Forms.Label
        $TitleLabel.Text = "🎬 Top 50 Movies Being Rented"
        $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
        $TitleLabel.Location = New-Object System.Drawing.Point(20, 15)
        $TitleLabel.AutoSize = $true
        $MoviesForm.Controls.Add($TitleLabel)
    
        # Genre Label
        $GenreLabel = New-Object System.Windows.Forms.Label
        $GenreLabel.Text = "Genre:"
        $GenreLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $GenreLabel.Location = New-Object System.Drawing.Point(400, 18)
        $GenreLabel.AutoSize = $true
        $MoviesForm.Controls.Add($GenreLabel)
    
        # Genre Dropdown with iTunes genre IDs
        $GenreCombo = New-Object System.Windows.Forms.ComboBox
        $GenreCombo.Location = New-Object System.Drawing.Point(455, 15)
        $GenreCombo.Size = New-Object System.Drawing.Size(170, 25)
        $GenreCombo.DropDownStyle = "DropDownList"
        $GenreCombo.BackColor = $ColorPanel
        $GenreCombo.ForeColor = "White"
        $GenreCombo.FlatStyle = "Flat"
    
        # Genre items with name and iTunes ID
        $Genres = @(
            @{Name = "All Genres"; Id = "" },
            @{Name = "Action & Adventure"; Id = "4401" },
            @{Name = "Comedy"; Id = "4404" },
            @{Name = "Drama"; Id = "4406" },
            @{Name = "Horror"; Id = "4408" },
            @{Name = "Kids & Family"; Id = "4410" },
            @{Name = "Romance"; Id = "4415" },
            @{Name = "Sci-Fi & Fantasy"; Id = "4416" },
            @{Name = "Thriller"; Id = "4419" },
            @{Name = "Documentary"; Id = "4405" },
            @{Name = "Anime"; Id = "4402" },
            @{Name = "Classics"; Id = "4403" },
            @{Name = "Independent"; Id = "4409" },
            @{Name = "Musicals"; Id = "4414" },
            @{Name = "Sports"; Id = "4420" },
            @{Name = "Western"; Id = "4421" }
        )
    
        foreach ($genre in $Genres) {
            $GenreCombo.Items.Add($genre.Name) | Out-Null
        }
        $GenreCombo.SelectedIndex = 0
        $MoviesForm.Controls.Add($GenreCombo)
    
        $SourceLabel = New-Object System.Windows.Forms.Label
        $SourceLabel.Text = "Source: iTunes Store"
        $SourceLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $SourceLabel.ForeColor = [System.Drawing.Color]::Gray
        $SourceLabel.Location = New-Object System.Drawing.Point(20, 50)
        $SourceLabel.AutoSize = $true
        $MoviesForm.Controls.Add($SourceLabel)
    
        # Create ImageList for cover art
        $ImageList = New-Object System.Windows.Forms.ImageList
        $ImageList.ImageSize = New-Object System.Drawing.Size(80, 120)
        $ImageList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
    
        $MoviesList = New-Object System.Windows.Forms.ListView
        $MoviesList.Location = New-Object System.Drawing.Point(20, 80)
        $MoviesList.Size = New-Object System.Drawing.Size(595, 430)
        $MoviesList.View = "LargeIcon"
        $MoviesList.LargeImageList = $ImageList
        $MoviesList.BackColor = $ColorPanel
        $MoviesList.ForeColor = "White"
        $MoviesList.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $MoviesList.BorderStyle = "None"
        $MoviesList.CheckBoxes = $true
        $MoviesList.MultiSelect = $true
        $MoviesForm.Controls.Add($MoviesList)
    
        # Selection counter label
        $SelectionCountLabel = New-Object System.Windows.Forms.Label
        $SelectionCountLabel.Text = "0 movies selected"
        $SelectionCountLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $SelectionCountLabel.ForeColor = $ColorAccent
        $SelectionCountLabel.Location = New-Object System.Drawing.Point(20, 525)
        $SelectionCountLabel.AutoSize = $true
        $MoviesForm.Controls.Add($SelectionCountLabel)
    
        # Update selection count when items are checked
        $MoviesList.Add_ItemChecked({
                param($listSender, $e)
                $checkedCount = $listSender.CheckedItems.Count
                if ($checkedCount -eq 1) {
                    $SelectionCountLabel.Text = "1 movie selected"
                }
                else {
                    $SelectionCountLabel.Text = "$checkedCount movies selected"
                }
            }.GetNewClosure())
    
        # Context menu for right-click actions
        $MoviesContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
        $MoviesContextMenu.BackColor = $ColorPanel
        $MoviesContextMenu.ForeColor = "White"
    
        $MenuViewTrailer = $MoviesContextMenu.Items.Add("🎬 View Trailer on YouTube")
        $MenuViewTrailer.Add_Click({
                if ($MoviesList.SelectedItems.Count -gt 0) {
                    $selectedItem = $MoviesList.SelectedItems[0]
                    # Extract movie title (remove the rank number prefix like "1. ")
                    $movieTitle = $selectedItem.Text -replace "^\d+\.\s*", ""
                    $searchQuery = [System.Uri]::EscapeDataString("$movieTitle official trailer")
            
                    try {
                        # Fetch YouTube search page with User-Agent header
                        $searchUrl = "https://www.youtube.com/results?search_query=$searchQuery"
                        $response = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -TimeoutSec 10 -Headers @{"User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" }
                
                        # Extract first video ID using regex
                        $videoMatches = [regex]::Matches($response.Content, '"videoId":"([a-zA-Z0-9_-]{11})"')
                        if ($videoMatches.Count -gt 0) {
                            $videoId = $videoMatches[0].Groups[1].Value
                            $directUrl = "https://www.youtube.com/watch?v=$videoId"
                            Start-Process $directUrl
                        }
                        else {
                            # Fallback to search results if parsing fails
                            Start-Process $searchUrl
                        }
                    }
                    catch {
                        # Fallback to search results on error
                        Start-Process "https://www.youtube.com/results?search_query=$searchQuery"
                    }
                }
            }.GetNewClosure())
    
        $MoviesList.ContextMenuStrip = $MoviesContextMenu
    
        # Track last clicked index for Shift+Click range selection
        $script:LastClickedIndex = -1
    
        # Handle click to toggle checkbox (left-click) or select item (right-click)
        # Supports Shift+Click for range selection
        $MoviesList.Add_MouseClick({
                param($sender, $e)
                $hitTest = $sender.HitTest($e.X, $e.Y)
                if ($hitTest.Item) {
                    $currentIndex = $hitTest.Item.Index
                    
                    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                        # Check if Shift is held for range selection
                        $shiftHeld = [System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Shift
                        
                        if ($shiftHeld -and $script:LastClickedIndex -ge 0) {
                            # Range selection: check/uncheck all items between last click and current
                            $startIndex = [Math]::Min($script:LastClickedIndex, $currentIndex)
                            $endIndex = [Math]::Max($script:LastClickedIndex, $currentIndex)
                            
                            # Determine the target state (use the opposite of current item's state)
                            $targetState = -not $hitTest.Item.Checked
                            
                            for ($i = $startIndex; $i -le $endIndex; $i++) {
                                $sender.Items[$i].Checked = $targetState
                            }
                        }
                        else {
                            # Single click: toggle checkbox
                            $hitTest.Item.Checked = -not $hitTest.Item.Checked
                        }
                        
                        # Update last clicked index
                        $script:LastClickedIndex = $currentIndex
                    }
                    elseif ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                        # Select item on right-click for context menu
                        $sender.SelectedItems.Clear()
                        $hitTest.Item.Selected = $true
                    }
                }
            }.GetNewClosure())
    
        $LoadingLabel = New-Object System.Windows.Forms.Label
        $LoadingLabel.Text = "Loading movies with cover art..."
        $LoadingLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        $LoadingLabel.Location = New-Object System.Drawing.Point(200, 300)
        $LoadingLabel.AutoSize = $true
        $MoviesForm.Controls.Add($LoadingLabel)
    
        # Add to Radarr Button
        $RadarrBtn = New-Object System.Windows.Forms.Button
        $RadarrBtn.Text = "📥 Add to Radarr"
        $RadarrBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $RadarrBtn.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 165, 0)
        $RadarrBtn.ForeColor = "Black"
        $RadarrBtn.FlatStyle = "Flat"
        $RadarrBtn.Size = New-Object System.Drawing.Size(150, 35)
        $RadarrBtn.Location = New-Object System.Drawing.Point(200, 570)
        $MoviesForm.Controls.Add($RadarrBtn)
    
        $RadarrBtn.Add_Click({
                # Check if any movies are selected
                if ($MoviesList.CheckedItems.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("Please check at least one movie to add to Radarr.", "No Selection", "OK", "Warning") | Out-Null
                    return
                }
        
                # Load saved configuration
                $ConfigPath = Join-Path $env:APPDATA "RadarrConfig.json"
                $SavedConfig = @{
                    RadarrUrl  = "http://localhost:7878"
                    ApiKey     = ""
                    RootFolder = "/movies"
                }
                if (Test-Path $ConfigPath) {
                    try {
                        $SavedConfig = Get-Content $ConfigPath -Raw | ConvertFrom-Json
                    }
                    catch { }
                }
        
                # Prompt for Radarr configuration
                $ConfigForm = New-Object System.Windows.Forms.Form
                $ConfigForm.Text = "Radarr Configuration"
                $ConfigForm.Size = New-Object System.Drawing.Size(450, 400)
                $ConfigForm.StartPosition = "CenterScreen"
                $ConfigForm.FormBorderStyle = "FixedDialog"
                $ConfigForm.BackColor = $ColorBackground
                $ConfigForm.ForeColor = "White"
        
                $UrlLabel = New-Object System.Windows.Forms.Label
                $UrlLabel.Text = "Radarr URL (e.g., http://localhost:7878):"
                $UrlLabel.Location = New-Object System.Drawing.Point(20, 20)
                $UrlLabel.AutoSize = $true
                $ConfigForm.Controls.Add($UrlLabel)
        
                $UrlBox = New-Object System.Windows.Forms.TextBox
                $UrlBox.Location = New-Object System.Drawing.Point(20, 45)
                $UrlBox.Size = New-Object System.Drawing.Size(390, 25)
                $UrlBox.Text = $SavedConfig.RadarrUrl
                $ConfigForm.Controls.Add($UrlBox)
        
                $ApiLabel = New-Object System.Windows.Forms.Label
                $ApiLabel.Text = "API Key:"
                $ApiLabel.Location = New-Object System.Drawing.Point(20, 80)
                $ApiLabel.AutoSize = $true
                $ConfigForm.Controls.Add($ApiLabel)
        
                $ApiBox = New-Object System.Windows.Forms.TextBox
                $ApiBox.Location = New-Object System.Drawing.Point(20, 105)
                $ApiBox.Size = New-Object System.Drawing.Size(300, 25)
                $ApiBox.Text = $SavedConfig.ApiKey
                $ConfigForm.Controls.Add($ApiBox)
        
                $ConnectBtn = New-Object System.Windows.Forms.Button
                $ConnectBtn.Text = "Connect"
                $ConnectBtn.Location = New-Object System.Drawing.Point(330, 103)
                $ConnectBtn.Size = New-Object System.Drawing.Size(80, 29)
                $ConnectBtn.BackColor = $ColorAccent
                $ConnectBtn.ForeColor = "White"
                $ConnectBtn.FlatStyle = "Flat"
                $ConfigForm.Controls.Add($ConnectBtn)
        
                $QualityLabel = New-Object System.Windows.Forms.Label
                $QualityLabel.Text = "Quality Profile:"
                $QualityLabel.Location = New-Object System.Drawing.Point(20, 140)
                $QualityLabel.AutoSize = $true
                $ConfigForm.Controls.Add($QualityLabel)
        
                $QualityCombo = New-Object System.Windows.Forms.ComboBox
                $QualityCombo.Location = New-Object System.Drawing.Point(20, 165)
                $QualityCombo.Size = New-Object System.Drawing.Size(390, 25)
                $QualityCombo.DropDownStyle = "DropDownList"
                $ConfigForm.Controls.Add($QualityCombo)
        
                $RootLabel = New-Object System.Windows.Forms.Label
                $RootLabel.Text = "Root Folder Path (e.g., /movies or C:\Movies):"
                $RootLabel.Location = New-Object System.Drawing.Point(20, 200)
                $RootLabel.AutoSize = $true
                $ConfigForm.Controls.Add($RootLabel)
        
                $RootBox = New-Object System.Windows.Forms.TextBox
                $RootBox.Location = New-Object System.Drawing.Point(20, 225)
                $RootBox.Size = New-Object System.Drawing.Size(390, 25)
                $RootBox.Text = $SavedConfig.RootFolder
                $ConfigForm.Controls.Add($RootBox)
        
                $AddBtn = New-Object System.Windows.Forms.Button
                $AddBtn.Text = "Add Movies"
                $AddBtn.Size = New-Object System.Drawing.Size(100, 30)
                $AddBtn.Location = New-Object System.Drawing.Point(200, 310)
                $AddBtn.BackColor = [System.Drawing.Color]::FromArgb(255, 76, 175, 80)
                $AddBtn.ForeColor = "White"
                $AddBtn.FlatStyle = "Flat"
                $AddBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $ConfigForm.Controls.Add($AddBtn)
        
                $CancelBtn = New-Object System.Windows.Forms.Button
                $CancelBtn.Text = "Cancel"
                $CancelBtn.Size = New-Object System.Drawing.Size(80, 30)
                $CancelBtn.Location = New-Object System.Drawing.Point(310, 310)
                $CancelBtn.BackColor = $ColorPanel
                $CancelBtn.ForeColor = "White"
                $CancelBtn.FlatStyle = "Flat"
                $CancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $ConfigForm.Controls.Add($CancelBtn)
        
                # Function to fetch profiles
                $FetchProfiles = {
                    $RadarrUrl = $UrlBox.Text.TrimEnd('/')
                    $ApiKey = $ApiBox.Text
            
                    if (-not $ApiKey) { return }
            
                    try {
                        $headers = @{ "X-Api-Key" = $ApiKey }
                        $profilesUrl = "$RadarrUrl/api/v3/qualityprofile"
                        $profiles = Invoke-RestMethod -Uri $profilesUrl -Headers $headers -Method Get -TimeoutSec 5
                
                        $QualityCombo.Items.Clear()
                        $QualityCombo.DisplayMember = "Name"
                        $QualityCombo.ValueMember = "Id"
                
                        foreach ($profile in $profiles) {
                            $QualityCombo.Items.Add($profile) | Out-Null
                        }
                
                        if ($QualityCombo.Items.Count -gt 0) {
                            $QualityCombo.SelectedIndex = 0
                        }
                        [System.Windows.Forms.MessageBox]::Show("Connected! Found $($profiles.Count) profiles.", "Success", "OK", "Information") | Out-Null
                    }
                    catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to connect to Radarr: $_", "Connection Error", "OK", "Error") | Out-Null
                    }
                }
        
                $ConnectBtn.Add_Click({ & $FetchProfiles }.GetNewClosure())
        
                # Auto-fetch if config exists
                if ($SavedConfig.ApiKey) {
                    try {
                        # Silent fetch
                        $RadarrUrl = $UrlBox.Text.TrimEnd('/')
                        $ApiKey = $ApiBox.Text
                        $headers = @{ "X-Api-Key" = $ApiKey }
                        $profilesUrl = "$RadarrUrl/api/v3/qualityprofile"
                        $profiles = Invoke-RestMethod -Uri $profilesUrl -Headers $headers -Method Get -TimeoutSec 5 -ErrorAction Stop
                
                        $QualityCombo.Items.Clear()
                        $QualityCombo.DisplayMember = "Name"
                        $QualityCombo.ValueMember = "Id"
                        foreach ($profile in $profiles) { $QualityCombo.Items.Add($profile) | Out-Null }
                        if ($QualityCombo.Items.Count -gt 0) { $QualityCombo.SelectedIndex = 0 }
                    }
                    catch { }
                }
        
                if ($ConfigForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $RadarrUrl = $UrlBox.Text.TrimEnd('/')
                    $ApiKey = $ApiBox.Text
                    $RootFolder = $RootBox.Text
                    $SelectedProfile = $QualityCombo.SelectedItem
            
                    if (-not $SelectedProfile) {
                        [System.Windows.Forms.MessageBox]::Show("Please select a Quality Profile.", "Error", "OK", "Error") | Out-Null
                        return
                    }
            
                    # Save configuration for next time
                    $ConfigToSave = @{
                        RadarrUrl  = $RadarrUrl
                        ApiKey     = $ApiKey
                        RootFolder = $RootFolder
                    } | ConvertTo-Json
                    $ConfigToSave | Set-Content $ConfigPath -Force
            
                    if (-not $ApiKey) {
                        [System.Windows.Forms.MessageBox]::Show("API Key is required.", "Error", "OK", "Error") | Out-Null
                        return
                    }
            
                    $successCount = 0
                    $failCount = 0
                    $headers = @{
                        "X-Api-Key"    = $ApiKey
                        "Content-Type" = "application/json"
                    }
            
                    foreach ($item in $MoviesList.CheckedItems) {
                        $movieTitle = $item.Text -replace "^\d+\.\s*", ""
                        # Remove extra info like RT score if present (e.g. "1. 🍅 95%`nTitle")
                        if ($movieTitle -match "`n") {
                            $movieTitle = ($movieTitle -split "`n")[1]
                        }
                
                        try {
                            # Search for the movie in Radarr
                            $searchUrl = "$RadarrUrl/api/v3/movie/lookup?term=$([System.Uri]::EscapeDataString($movieTitle))"
                            $searchResult = Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Get -TimeoutSec 10
                    
                            if ($searchResult -and $searchResult.Count -gt 0) {
                                $movie = $searchResult[0]
                        
                                # Check if already in library
                                $existingUrl = "$RadarrUrl/api/v3/movie?tmdbId=$($movie.tmdbId)"
                                $existing = Invoke-RestMethod -Uri $existingUrl -Headers $headers -Method Get -TimeoutSec 5 -ErrorAction SilentlyContinue
                        
                                if ($existing -and $existing.Count -gt 0) {
                                    $failCount++
                                    continue
                                }
                        
                                # Add to Radarr
                                $addBody = @{
                                    title            = $movie.title
                                    tmdbId           = $movie.tmdbId
                                    year             = $movie.year
                                    qualityProfileId = $SelectedProfile.id
                                    rootFolderPath   = $RootFolder
                                    monitored        = $true
                                    addOptions       = @{
                                        searchForMovie = $true
                                    }
                                } | ConvertTo-Json -Depth 10
                        
                                $addUrl = "$RadarrUrl/api/v3/movie"
                                Invoke-RestMethod -Uri $addUrl -Headers $headers -Method Post -Body $addBody -TimeoutSec 10 | Out-Null
                                $successCount++
                            }
                            else {
                                $failCount++
                            }
                        }
                        catch {
                            $failCount++
                        }
                    }
            
                    [System.Windows.Forms.MessageBox]::Show("Added $successCount movie(s) to Radarr.`n$failCount movie(s) failed or already exist.", "Radarr Import Complete", "OK", "Information") | Out-Null
                }
            }.GetNewClosure())
    
        $CloseBtn = New-Object System.Windows.Forms.Button
        $CloseBtn.Text = "Close"
        $CloseBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $CloseBtn.BackColor = $ColorPanel
        $CloseBtn.ForeColor = "White"
        $CloseBtn.FlatStyle = "Flat"
        $CloseBtn.Size = New-Object System.Drawing.Size(100, 35)
        $CloseBtn.Location = New-Object System.Drawing.Point(530, 570)
        $MoviesForm.Controls.Add($CloseBtn)
        $CloseBtn.Add_Click({ $this.FindForm().Close() }.GetNewClosure())
    
        # Function to load movies based on genre
        $LoadMovies = {
            param($GenreId)
        
            $MoviesList.Items.Clear()
            $ImageList.Images.Clear()
            $LoadingLabel.Visible = $true
            $LoadingLabel.Text = "Loading movies with cover art..."
            $LoadingLabel.ForeColor = "White"
            $SelectionCountLabel.Text = "0 movies selected"
            [System.Windows.Forms.Application]::DoEvents()
        
            try {
                # Build URL with optional genre filter
                if ($GenreId -and $GenreId -ne "") {
                    $FeedUrl = "https://itunes.apple.com/us/rss/topmovies/genre=$GenreId/limit=50/xml"
                    $genreName = ($Genres | Where-Object { $_.Id -eq $GenreId }).Name
                    $SourceLabel.Text = "Source: iTunes Store - $genreName"
                }
                else {
                    $FeedUrl = "https://itunes.apple.com/us/rss/topmovies/limit=50/xml"
                    $SourceLabel.Text = "Source: iTunes Store - All Genres"
                }
            
                [xml]$rss = (Invoke-WebRequest -UseBasicParsing -Uri $FeedUrl -TimeoutSec 15).Content
            
                $LoadingLabel.Visible = $false
            
                $Rank = 0
                $WebClient = New-Object System.Net.WebClient
            
                foreach ($entry in $rss.feed.entry) {
                    $Rank++
                    $rawTitle = [string]$entry.title
                    $title = ($rawTitle -split " - ")[0].Trim()
                
                    # Try to get Rotten Tomatoes score from OMDb API
                    $rtScore = ""
                    $rtIcon = ""
                    try {
                        $omdbUrl = "https://www.omdbapi.com/?t=$([System.Uri]::EscapeDataString($title))&apikey=b6003d8a"
                        $omdbResult = Invoke-RestMethod -Uri $omdbUrl -TimeoutSec 3 -ErrorAction SilentlyContinue
                        if ($omdbResult -and $omdbResult.Ratings) {
                            $rtRating = $omdbResult.Ratings | Where-Object { $_.Source -eq "Rotten Tomatoes" } | Select-Object -First 1
                            if ($rtRating) {
                                $rtScore = $rtRating.Value
                                $rtPercent = [int]($rtScore -replace '%', '')
                                $rtIcon = if ($rtPercent -ge 60) { "🍅" } else { "🤢" }
                            }
                        }
                    }
                    catch { }
                
                    # Get the largest image (170 height)
                    $imageUrl = ($entry.ChildNodes | Where-Object { $_.LocalName -eq "image" -and $_.height -eq "170" } | Select-Object -First 1)."#text"
                    if (-not $imageUrl) {
                        $imageUrl = ($entry.ChildNodes | Where-Object { $_.LocalName -eq "image" } | Select-Object -Last 1)."#text"
                    }
                
                    # Download and add image
                    try {
                        if ($imageUrl) {
                            $imageData = $WebClient.DownloadData($imageUrl)
                            $memStream = New-Object System.IO.MemoryStream(, $imageData)
                            $img = [System.Drawing.Image]::FromStream($memStream)
                            $ImageList.Images.Add($Rank.ToString(), $img)
                        }
                    }
                    catch { }
                
                    # Build display text with RT score
                    $displayText = "$Rank. $title"
                    if ($rtScore) {
                        $displayText = "$Rank. $rtIcon $rtScore`n$title"
                    }
                    $Item = New-Object System.Windows.Forms.ListViewItem($displayText, $Rank.ToString())
                    $MoviesList.Items.Add($Item) | Out-Null
                
                    # Update UI every 5 movies
                    if ($Rank % 5 -eq 0) {
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }
            
                $WebClient.Dispose()
            }
            catch {
                $LoadingLabel.Text = "Failed to fetch movies.`nCheck internet connection."
                $LoadingLabel.ForeColor = [System.Drawing.Color]::Salmon
            }
        }
    
        # Genre selection change handler - reload movies when genre changes  
        $GenreCombo.Add_SelectedIndexChanged({
                $selectedGenreIndex = $GenreCombo.SelectedIndex
                $selectedGenreId = $Genres[$selectedGenreIndex].Id
        
                # Clear existing items
                $MoviesList.Items.Clear()
                $ImageList.Images.Clear()
                $LoadingLabel.Visible = $true
                $LoadingLabel.Text = "Loading movies..."
                $LoadingLabel.ForeColor = "White"
                $SelectionCountLabel.Text = "0 movies selected"
                [System.Windows.Forms.Application]::DoEvents()
        
                try {
                    # Build URL with optional genre filter
                    if ($selectedGenreId -and $selectedGenreId -ne "") {
                        $FeedUrl = "https://itunes.apple.com/us/rss/topmovies/genre=$selectedGenreId/limit=50/xml"
                        $genreName = $Genres[$selectedGenreIndex].Name
                        $SourceLabel.Text = "Source: iTunes Store - $genreName"
                    }
                    else {
                        $FeedUrl = "https://itunes.apple.com/us/rss/topmovies/limit=50/xml"
                        $SourceLabel.Text = "Source: iTunes Store - All Genres"
                    }
                
                    [xml]$rss = (Invoke-WebRequest -UseBasicParsing -Uri $FeedUrl -TimeoutSec 15).Content
                
                    $LoadingLabel.Visible = $false
                
                    $Rank = 0
                    $WebClient = New-Object System.Net.WebClient
                
                    foreach ($entry in $rss.feed.entry) {
                        $Rank++
                        $rawTitle = [string]$entry.title
                        $title = ($rawTitle -split " - ")[0].Trim()
                    
                        # Try to get Rotten Tomatoes score from OMDb API
                        $rtScore = ""
                        $rtIcon = ""
                        try {
                            $omdbUrl = "https://www.omdbapi.com/?t=$([System.Uri]::EscapeDataString($title))&apikey=b6003d8a"
                            $omdbResult = Invoke-RestMethod -Uri $omdbUrl -TimeoutSec 3 -ErrorAction SilentlyContinue
                            if ($omdbResult -and $omdbResult.Ratings) {
                                $rtRating = $omdbResult.Ratings | Where-Object { $_.Source -eq "Rotten Tomatoes" } | Select-Object -First 1
                                if ($rtRating) {
                                    $rtScore = $rtRating.Value
                                    $rtPercent = [int]($rtScore -replace '%', '')
                                    $rtIcon = if ($rtPercent -ge 60) { "🍅" } else { "🤢" }
                                }
                            }
                        }
                        catch { }
                    
                        # Get the largest image (170 height)
                        $imageUrl = ($entry.ChildNodes | Where-Object { $_.LocalName -eq "image" -and $_.height -eq "170" } | Select-Object -First 1)."#text"
                        if (-not $imageUrl) {
                            $imageUrl = ($entry.ChildNodes | Where-Object { $_.LocalName -eq "image" } | Select-Object -Last 1)."#text"
                        }
                    
                        # Download and add image
                        try {
                            if ($imageUrl) {
                                $imageData = $WebClient.DownloadData($imageUrl)
                                $memStream = New-Object System.IO.MemoryStream(, $imageData)
                                $img = [System.Drawing.Image]::FromStream($memStream)
                                $ImageList.Images.Add($Rank.ToString(), $img)
                            }
                        }
                        catch { }
                    
                        # Build display text with RT score
                        $displayText = "$Rank. $title"
                        if ($rtScore) {
                            $displayText = "$Rank. $rtIcon $rtScore`n$title"
                        }
                        $Item = New-Object System.Windows.Forms.ListViewItem($displayText, $Rank.ToString())
                        $MoviesList.Items.Add($Item) | Out-Null
                
                        # Update UI every 5 movies
                        if ($Rank % 5 -eq 0) {
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                    }
            
                    $WebClient.Dispose()
                }
                catch {
                    $LoadingLabel.Text = "Failed to fetch movies."
                    $LoadingLabel.ForeColor = [System.Drawing.Color]::Salmon
                }
            }.GetNewClosure())
    
        # Show form and load initial movies
        $MoviesForm.Show()
        $MoviesForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        # Load initial movies list (All Genres by default)
        & $LoadMovies ""
    })

# Letterboxd Import Button Click - Import movies from a Letterboxd list
$LetterboxdButton.Add_Click({
        # Create URL input dialog
        $UrlForm = New-Object System.Windows.Forms.Form
        $UrlForm.Text = "Import Letterboxd List"
        $UrlForm.Size = New-Object System.Drawing.Size(550, 180)
        $UrlForm.StartPosition = "CenterScreen"
        $UrlForm.FormBorderStyle = "FixedDialog"
        $UrlForm.MaximizeBox = $false
        $UrlForm.MinimizeBox = $false
        $UrlForm.BackColor = $ColorBackground
        $UrlForm.ForeColor = "White"
        
        $UrlLabel = New-Object System.Windows.Forms.Label
        $UrlLabel.Text = "Enter Letterboxd List URL:"
        $UrlLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $UrlLabel.Location = New-Object System.Drawing.Point(20, 20)
        $UrlLabel.AutoSize = $true
        $UrlForm.Controls.Add($UrlLabel)
        
        $UrlTextBox = New-Object System.Windows.Forms.TextBox
        $UrlTextBox.Location = New-Object System.Drawing.Point(20, 50)
        $UrlTextBox.Size = New-Object System.Drawing.Size(490, 30)
        $UrlTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $UrlTextBox.BackColor = $ColorPanel
        $UrlTextBox.ForeColor = "White"
        $UrlTextBox.Text = "https://letterboxd.com/"
        $UrlForm.Controls.Add($UrlTextBox)
        
        $ImportBtn = New-Object System.Windows.Forms.Button
        $ImportBtn.Text = "Import"
        $ImportBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $ImportBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 136)
        $ImportBtn.ForeColor = "White"
        $ImportBtn.FlatStyle = "Flat"
        $ImportBtn.FlatAppearance.BorderSize = 0
        $ImportBtn.Size = New-Object System.Drawing.Size(100, 35)
        $ImportBtn.Location = New-Object System.Drawing.Point(300, 95)
        $ImportBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $UrlForm.Controls.Add($ImportBtn)
        
        $CancelBtn = New-Object System.Windows.Forms.Button
        $CancelBtn.Text = "Cancel"
        $CancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $CancelBtn.BackColor = $ColorPanel
        $CancelBtn.ForeColor = "White"
        $CancelBtn.FlatStyle = "Flat"
        $CancelBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
        $CancelBtn.Size = New-Object System.Drawing.Size(100, 35)
        $CancelBtn.Location = New-Object System.Drawing.Point(410, 95)
        $CancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $UrlForm.Controls.Add($CancelBtn)
        
        $UrlForm.AcceptButton = $ImportBtn
        $UrlForm.CancelButton = $CancelBtn
        
        if ($UrlForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }
        
        $ListUrl = $UrlTextBox.Text.Trim()
        $UrlForm.Dispose()
        
        if (-not $ListUrl -or $ListUrl -eq "https://letterboxd.com/") {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid Letterboxd list URL.", "Error", "OK", "Warning") | Out-Null
            return
        }
        
        # Create the movies popup form
        $MoviesForm = New-Object System.Windows.Forms.Form
        $MoviesForm.Text = "Letterboxd Import"
        $MoviesForm.Size = New-Object System.Drawing.Size(650, 650)
        $MoviesForm.StartPosition = "CenterScreen"
        $MoviesForm.FormBorderStyle = "FixedDialog"
        $MoviesForm.MaximizeBox = $false
        $MoviesForm.BackColor = $ColorBackground
        $MoviesForm.ForeColor = "White"
        
        $TitleLabel = New-Object System.Windows.Forms.Label
        $TitleLabel.Text = "📋 Letterboxd List Import"
        $TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
        $TitleLabel.Location = New-Object System.Drawing.Point(20, 15)
        $TitleLabel.AutoSize = $true
        $MoviesForm.Controls.Add($TitleLabel)
        
        $SourceLabel = New-Object System.Windows.Forms.Label
        $SourceLabel.Text = "Source: $ListUrl"
        $SourceLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $SourceLabel.ForeColor = [System.Drawing.Color]::Gray
        $SourceLabel.Location = New-Object System.Drawing.Point(20, 50)
        $SourceLabel.Size = New-Object System.Drawing.Size(600, 20)
        $MoviesForm.Controls.Add($SourceLabel)
        
        # Create ImageList for cover art
        $ImageList = New-Object System.Windows.Forms.ImageList
        $ImageList.ImageSize = New-Object System.Drawing.Size(80, 120)
        $ImageList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
        
        $MoviesList = New-Object System.Windows.Forms.ListView
        $MoviesList.Location = New-Object System.Drawing.Point(20, 80)
        $MoviesList.Size = New-Object System.Drawing.Size(595, 430)
        $MoviesList.View = "LargeIcon"
        $MoviesList.LargeImageList = $ImageList
        $MoviesList.BackColor = $ColorPanel
        $MoviesList.ForeColor = "White"
        $MoviesList.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $MoviesList.BorderStyle = "None"
        $MoviesList.CheckBoxes = $true
        $MoviesList.MultiSelect = $true
        $MoviesForm.Controls.Add($MoviesList)
        
        # Create context menu for right-click actions
        $ContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
        $ContextMenu.BackColor = $ColorPanel
        $ContextMenu.ForeColor = "White"
        
        $ViewTrailerMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $ViewTrailerMenuItem.Text = "🎬 View Trailer on YouTube"
        $ViewTrailerMenuItem.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $ContextMenu.Items.Add($ViewTrailerMenuItem) | Out-Null
        
        $ViewTrailerMenuItem.Add_Click({
                # Get the item that was right-clicked
                $selectedItem = $MoviesList.FocusedItem
                
                if (-not $selectedItem) {
                    [System.Windows.Forms.MessageBox]::Show("Please right-click on a movie to view its trailer.", "No Selection", "OK", "Warning") | Out-Null
                    return
                }
                
                # Extract title from the item text (format: "1. Title\n(Year)")
                $itemText = $selectedItem.Text
                $titleMatch = [regex]::Match($itemText, '^\d+\.\s*(.+?)(?:\n|\r|$)')
                if ($titleMatch.Success) {
                    $movieTitle = $titleMatch.Groups[1].Value.Trim()
                    
                    # Extract year if present
                    $yearMatch = [regex]::Match($itemText, '\((\d{4})\)')
                    $searchQuery = if ($yearMatch.Success) { 
                        "$movieTitle $($yearMatch.Groups[1].Value) official trailer"
                    }
                    else {
                        "$movieTitle official trailer"
                    }
                    
                    # Open YouTube search
                    $youtubeUrl = "https://www.youtube.com/results?search_query=$([System.Uri]::EscapeDataString($searchQuery))"
                    Start-Process $youtubeUrl
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("Could not parse movie title.", "Error", "OK", "Error") | Out-Null
                }
            }.GetNewClosure())
        
        $MoviesList.ContextMenuStrip = $ContextMenu
        
        # Selection counter label
        $SelectionCountLabel = New-Object System.Windows.Forms.Label
        $SelectionCountLabel.Text = "0 movies selected"
        $SelectionCountLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $SelectionCountLabel.ForeColor = $ColorAccent
        $SelectionCountLabel.Location = New-Object System.Drawing.Point(20, 525)
        $SelectionCountLabel.AutoSize = $true
        $MoviesForm.Controls.Add($SelectionCountLabel)
        
        # Update selection count when items are checked
        $MoviesList.Add_ItemChecked({
                param($listSender, $e)
                $checkedCount = $listSender.CheckedItems.Count
                if ($checkedCount -eq 1) {
                    $SelectionCountLabel.Text = "1 movie selected"
                }
                else {
                    $SelectionCountLabel.Text = "$checkedCount movies selected"
                }
            }.GetNewClosure())
        
        # Handle click to toggle checkbox
        $MoviesList.Add_MouseClick({
                param($sender, $e)
                $hitTest = $sender.HitTest($e.X, $e.Y)
                if ($hitTest.Item -and $e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    $hitTest.Item.Checked = -not $hitTest.Item.Checked
                }
            }.GetNewClosure())
        
        $LoadingLabel = New-Object System.Windows.Forms.Label
        $LoadingLabel.Text = "Loading movies from Letterboxd..."
        $LoadingLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        $LoadingLabel.Location = New-Object System.Drawing.Point(200, 300)
        $LoadingLabel.AutoSize = $true
        $MoviesForm.Controls.Add($LoadingLabel)
        
        # Add to Radarr Button
        $RadarrBtn = New-Object System.Windows.Forms.Button
        $RadarrBtn.Text = "📥 Add to Radarr"
        $RadarrBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $RadarrBtn.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 165, 0)
        $RadarrBtn.ForeColor = "Black"
        $RadarrBtn.FlatStyle = "Flat"
        $RadarrBtn.Size = New-Object System.Drawing.Size(150, 35)
        $RadarrBtn.Location = New-Object System.Drawing.Point(200, 570)
        $MoviesForm.Controls.Add($RadarrBtn)
        
        $RadarrBtn.Add_Click({
                # Check if any movies are selected
                if ($MoviesList.CheckedItems.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("Please check at least one movie to add to Radarr.", "No Selection", "OK", "Warning") | Out-Null
                    return
                }
                
                # Load saved configuration
                $ConfigPath = Join-Path $env:APPDATA "RadarrConfig.json"
                $SavedConfig = @{
                    RadarrUrl  = "http://localhost:7878"
                    ApiKey     = ""
                    RootFolder = "/movies"
                }
                if (Test-Path $ConfigPath) {
                    try {
                        $SavedConfig = Get-Content $ConfigPath -Raw | ConvertFrom-Json
                    }
                    catch { }
                }
                
                # Prompt for Radarr configuration
                $ConfigForm = New-Object System.Windows.Forms.Form
                $ConfigForm.Text = "Radarr Configuration"
                $ConfigForm.Size = New-Object System.Drawing.Size(450, 400)
                $ConfigForm.StartPosition = "CenterScreen"
                $ConfigForm.FormBorderStyle = "FixedDialog"
                $ConfigForm.BackColor = $ColorBackground
                $ConfigForm.ForeColor = "White"
                
                $UrlLabel = New-Object System.Windows.Forms.Label
                $UrlLabel.Text = "Radarr URL (e.g., http://localhost:7878):"
                $UrlLabel.Location = New-Object System.Drawing.Point(20, 20)
                $UrlLabel.AutoSize = $true
                $ConfigForm.Controls.Add($UrlLabel)
                
                $UrlBox = New-Object System.Windows.Forms.TextBox
                $UrlBox.Location = New-Object System.Drawing.Point(20, 45)
                $UrlBox.Size = New-Object System.Drawing.Size(390, 25)
                $UrlBox.Text = $SavedConfig.RadarrUrl
                $ConfigForm.Controls.Add($UrlBox)
                
                $ApiLabel = New-Object System.Windows.Forms.Label
                $ApiLabel.Text = "API Key:"
                $ApiLabel.Location = New-Object System.Drawing.Point(20, 80)
                $ApiLabel.AutoSize = $true
                $ConfigForm.Controls.Add($ApiLabel)
                
                $ApiBox = New-Object System.Windows.Forms.TextBox
                $ApiBox.Location = New-Object System.Drawing.Point(20, 105)
                $ApiBox.Size = New-Object System.Drawing.Size(300, 25)
                $ApiBox.Text = $SavedConfig.ApiKey
                $ConfigForm.Controls.Add($ApiBox)
                
                $ConnectBtn = New-Object System.Windows.Forms.Button
                $ConnectBtn.Text = "Connect"
                $ConnectBtn.Location = New-Object System.Drawing.Point(330, 103)
                $ConnectBtn.Size = New-Object System.Drawing.Size(80, 29)
                $ConnectBtn.BackColor = $ColorAccent
                $ConnectBtn.ForeColor = "White"
                $ConnectBtn.FlatStyle = "Flat"
                $ConfigForm.Controls.Add($ConnectBtn)
                
                $QualityLabel = New-Object System.Windows.Forms.Label
                $QualityLabel.Text = "Quality Profile:"
                $QualityLabel.Location = New-Object System.Drawing.Point(20, 140)
                $QualityLabel.AutoSize = $true
                $ConfigForm.Controls.Add($QualityLabel)
                
                $QualityCombo = New-Object System.Windows.Forms.ComboBox
                $QualityCombo.Location = New-Object System.Drawing.Point(20, 165)
                $QualityCombo.Size = New-Object System.Drawing.Size(390, 25)
                $QualityCombo.DropDownStyle = "DropDownList"
                $ConfigForm.Controls.Add($QualityCombo)
                
                $RootLabel = New-Object System.Windows.Forms.Label
                $RootLabel.Text = "Root Folder Path (e.g., /movies or C:\Movies):"
                $RootLabel.Location = New-Object System.Drawing.Point(20, 200)
                $RootLabel.AutoSize = $true
                $ConfigForm.Controls.Add($RootLabel)
                
                $RootBox = New-Object System.Windows.Forms.TextBox
                $RootBox.Location = New-Object System.Drawing.Point(20, 225)
                $RootBox.Size = New-Object System.Drawing.Size(390, 25)
                $RootBox.Text = $SavedConfig.RootFolder
                $ConfigForm.Controls.Add($RootBox)
                
                $AddBtn = New-Object System.Windows.Forms.Button
                $AddBtn.Text = "Add Movies"
                $AddBtn.Size = New-Object System.Drawing.Size(100, 30)
                $AddBtn.Location = New-Object System.Drawing.Point(200, 310)
                $AddBtn.BackColor = [System.Drawing.Color]::FromArgb(255, 76, 175, 80)
                $AddBtn.ForeColor = "White"
                $AddBtn.FlatStyle = "Flat"
                $AddBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $ConfigForm.Controls.Add($AddBtn)
                
                $CancelConfigBtn = New-Object System.Windows.Forms.Button
                $CancelConfigBtn.Text = "Cancel"
                $CancelConfigBtn.Size = New-Object System.Drawing.Size(80, 30)
                $CancelConfigBtn.Location = New-Object System.Drawing.Point(310, 310)
                $CancelConfigBtn.BackColor = $ColorPanel
                $CancelConfigBtn.ForeColor = "White"
                $CancelConfigBtn.FlatStyle = "Flat"
                $CancelConfigBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $ConfigForm.Controls.Add($CancelConfigBtn)
                
                # Function to fetch profiles
                $FetchProfiles = {
                    $RadarrUrl = $UrlBox.Text.TrimEnd('/')
                    $ApiKey = $ApiBox.Text
                    
                    if (-not $ApiKey) { return }
                    
                    try {
                        $headers = @{ "X-Api-Key" = $ApiKey }
                        $profilesUrl = "$RadarrUrl/api/v3/qualityprofile"
                        $profiles = Invoke-RestMethod -Uri $profilesUrl -Headers $headers -Method Get -TimeoutSec 5
                        
                        $QualityCombo.Items.Clear()
                        $QualityCombo.DisplayMember = "Name"
                        $QualityCombo.ValueMember = "Id"
                        
                        foreach ($profile in $profiles) {
                            $QualityCombo.Items.Add($profile) | Out-Null
                        }
                        
                        if ($QualityCombo.Items.Count -gt 0) {
                            $QualityCombo.SelectedIndex = 0
                        }
                        [System.Windows.Forms.MessageBox]::Show("Connected! Found $($profiles.Count) profiles.", "Success", "OK", "Information") | Out-Null
                    }
                    catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to connect to Radarr: $_", "Connection Error", "OK", "Error") | Out-Null
                    }
                }
                
                $ConnectBtn.Add_Click({ & $FetchProfiles }.GetNewClosure())
                
                # Auto-fetch if config exists
                if ($SavedConfig.ApiKey) {
                    try {
                        $RadarrUrl = $UrlBox.Text.TrimEnd('/')
                        $ApiKey = $ApiBox.Text
                        $headers = @{ "X-Api-Key" = $ApiKey }
                        $profilesUrl = "$RadarrUrl/api/v3/qualityprofile"
                        $profiles = Invoke-RestMethod -Uri $profilesUrl -Headers $headers -Method Get -TimeoutSec 5 -ErrorAction Stop
                        
                        $QualityCombo.Items.Clear()
                        $QualityCombo.DisplayMember = "Name"
                        $QualityCombo.ValueMember = "Id"
                        foreach ($profile in $profiles) { $QualityCombo.Items.Add($profile) | Out-Null }
                        if ($QualityCombo.Items.Count -gt 0) { $QualityCombo.SelectedIndex = 0 }
                    }
                    catch { }
                }
                
                if ($ConfigForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $RadarrUrl = $UrlBox.Text.TrimEnd('/')
                    $ApiKey = $ApiBox.Text
                    $RootFolder = $RootBox.Text
                    $SelectedProfile = $QualityCombo.SelectedItem
                    
                    if (-not $SelectedProfile) {
                        [System.Windows.Forms.MessageBox]::Show("Please select a Quality Profile.", "Error", "OK", "Error") | Out-Null
                        return
                    }
                    
                    # Save configuration for next time
                    $ConfigToSave = @{
                        RadarrUrl  = $RadarrUrl
                        ApiKey     = $ApiKey
                        RootFolder = $RootFolder
                    } | ConvertTo-Json
                    $ConfigToSave | Set-Content $ConfigPath -Force
                    
                    if (-not $ApiKey) {
                        [System.Windows.Forms.MessageBox]::Show("API Key is required.", "Error", "OK", "Error") | Out-Null
                        return
                    }
                    
                    $successCount = 0
                    $failCount = 0
                    $headers = @{
                        "X-Api-Key"    = $ApiKey
                        "Content-Type" = "application/json"
                    }
                    
                    foreach ($item in $MoviesList.CheckedItems) {
                        $movieTitle = $item.Text -replace "^\d+\.\s*", ""
                        
                        try {
                            # Search for the movie in Radarr
                            $searchUrl = "$RadarrUrl/api/v3/movie/lookup?term=$([System.Uri]::EscapeDataString($movieTitle))"
                            $searchResult = Invoke-RestMethod -Uri $searchUrl -Headers $headers -Method Get -TimeoutSec 10
                            
                            if ($searchResult -and $searchResult.Count -gt 0) {
                                $movie = $searchResult[0]
                                
                                # Check if already in library
                                $existingUrl = "$RadarrUrl/api/v3/movie?tmdbId=$($movie.tmdbId)"
                                $existing = Invoke-RestMethod -Uri $existingUrl -Headers $headers -Method Get -TimeoutSec 5 -ErrorAction SilentlyContinue
                                
                                if ($existing -and $existing.Count -gt 0) {
                                    $failCount++
                                    continue
                                }
                                
                                # Add to Radarr
                                $addBody = @{
                                    title            = $movie.title
                                    tmdbId           = $movie.tmdbId
                                    year             = $movie.year
                                    qualityProfileId = $SelectedProfile.id
                                    rootFolderPath   = $RootFolder
                                    monitored        = $true
                                    addOptions       = @{
                                        searchForMovie = $true
                                    }
                                } | ConvertTo-Json -Depth 10
                                
                                $addUrl = "$RadarrUrl/api/v3/movie"
                                Invoke-RestMethod -Uri $addUrl -Headers $headers -Method Post -Body $addBody -TimeoutSec 10 | Out-Null
                                $successCount++
                            }
                            else {
                                $failCount++
                            }
                        }
                        catch {
                            $failCount++
                        }
                    }
                    
                    [System.Windows.Forms.MessageBox]::Show("Added $successCount movie(s) to Radarr.`n$failCount movie(s) failed or already exist.", "Radarr Import Complete", "OK", "Information") | Out-Null
                }
            }.GetNewClosure())
        
        $CloseBtn = New-Object System.Windows.Forms.Button
        $CloseBtn.Text = "Close"
        $CloseBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $CloseBtn.BackColor = $ColorPanel
        $CloseBtn.ForeColor = "White"
        $CloseBtn.FlatStyle = "Flat"
        $CloseBtn.Size = New-Object System.Drawing.Size(100, 35)
        $CloseBtn.Location = New-Object System.Drawing.Point(515, 570)
        $MoviesForm.Controls.Add($CloseBtn)
        $CloseBtn.Add_Click({ $this.FindForm().Close() }.GetNewClosure())
        
        # Show the form
        $MoviesForm.Show()
        $MoviesForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        # Fetch and parse movies from Letterboxd
        try {
            $allMovies = @()
            $pageNum = 1
            $baseUrl = $ListUrl.TrimEnd('/')
            
            # Fetch up to 3 pages (Letterboxd shows ~100 movies per page)
            while ($pageNum -le 3) {
                $pageUrl = if ($pageNum -eq 1) { $baseUrl } else { "$baseUrl/page/$pageNum/" }
                
                try {
                    $response = Invoke-WebRequest -Uri $pageUrl -UseBasicParsing -TimeoutSec 15 -Headers @{
                        "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                    }
                    
                    # Method 1: Extract from data-item-name attribute (most reliable - contains "Title (Year)")
                    $itemNameMatches = [regex]::Matches($response.Content, 'data-item-name="([^"]+)"')
                    foreach ($m in $itemNameMatches) {
                        $text = $m.Groups[1].Value.Trim()
                        if ($text -match '^(.+) \((\d{4})\)$') {
                            $title = $Matches[1]
                            $year = $Matches[2]
                            if (-not ($allMovies | Where-Object { $_.Title -eq $title -and $_.Year -eq $year })) {
                                $allMovies += @{ Title = $title; Year = $year; Slug = "" }
                            }
                        }
                    }
                    
                    # Method 2: Extract from /film/slug/ href pattern (fallback)
                    $hrefMatches = [regex]::Matches($response.Content, 'href="/film/([^/"]+)/"')
                    foreach ($m in $hrefMatches) {
                        $slug = $m.Groups[1].Value
                        # Convert slug to title: the-shawshank-redemption -> The Shawshank Redemption
                        $titleFromSlug = ($slug -replace '-', ' ') -replace '(\b\w)', { $_.Groups[0].Value.ToUpper() }
                        if ($titleFromSlug -and -not ($allMovies | Where-Object { $_.Slug -eq $slug -or $_.Title -eq $titleFromSlug })) {
                            $allMovies += @{ Title = $titleFromSlug; Year = ""; Slug = $slug }
                        }
                    }
                    
                    # Method 3: Also try frame-title spans (Title (Year) format) 
                    $frameTitleMatches = [regex]::Matches($response.Content, 'class="frame-title"[^>]*>([^<]+)<')
                    foreach ($m in $frameTitleMatches) {
                        $text = $m.Groups[1].Value.Trim()
                        if ($text -match '^(.+) \((\d{4})\)$') {
                            $title = $Matches[1]
                            $year = $Matches[2]
                            # Update year if we have a matching slug entry without year
                            $existing = $allMovies | Where-Object { $_.Title -eq $title -and (-not $_.Year -or $_.Year -eq "") }
                            if ($existing) {
                                $existing.Year = $year
                            }
                            elseif (-not ($allMovies | Where-Object { $_.Title -eq $title })) {
                                $allMovies += @{ Title = $title; Year = $year; Slug = "" }
                            }
                        }
                    }
                    
                    # Method 4: Try img alt text (for pages where JS has been cached/prerendered)
                    $altMatches = [regex]::Matches($response.Content, 'alt="Poster for ([^"]+) \((\d{4})\)"')
                    foreach ($m in $altMatches) {
                        $title = $m.Groups[1].Value
                        $year = $m.Groups[2].Value
                        if (-not ($allMovies | Where-Object { $_.Title -eq $title -and $_.Year -eq $year })) {
                            $allMovies += @{ Title = $title; Year = $year; Slug = "" }
                        }
                    }
                    
                    # Check if there's a next page link
                    if ($response.Content -notmatch "/page/$($pageNum + 1)/") {
                        break
                    }
                    $pageNum++
                }
                catch {
                    break
                }
            }
            
            $LoadingLabel.Visible = $false
            
            if ($allMovies.Count -eq 0) {
                $LoadingLabel.Text = "No movies found. Make sure the URL is a valid Letterboxd list."
                $LoadingLabel.ForeColor = [System.Drawing.Color]::Orange
                $LoadingLabel.Visible = $true
                return
            }
            
            $Rank = 0
            $WebClient = New-Object System.Net.WebClient
            $addedTitles = @{}  # Track unique titles
            
            foreach ($movie in $allMovies) {
                $title = $movie.Title
                $year = $movie.Year
                
                # Skip if no title
                if (-not $title) { continue }
                
                # Create unique key
                $fullTitle = if ($year) { "$title ($year)" } else { $title }
                
                # Skip duplicates
                if ($addedTitles.ContainsKey($fullTitle)) { continue }
                $addedTitles[$fullTitle] = $true
                
                $Rank++
                $displayText = if ($year) { "$Rank. $title`n($year)" } else { "$Rank. $title" }
                
                # Try to get poster from TMDB (more reliable than OMDb)
                $posterLoaded = $false
                try {
                    # Ensure TLS 1.2 is used
                    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                    
                    # TMDB API v3 - search for movie
                    $tmdbApiKey = "2dca580c2a14b55200e784d157207b4d"  # Public API key
                    $searchQuery = [System.Uri]::EscapeDataString($title)
                    $tmdbUrl = "https://api.themoviedb.org/3/search/movie?api_key=$tmdbApiKey&query=$searchQuery"
                    if ($year) { $tmdbUrl += "&year=$year" }
                    
                    $tmdbResult = Invoke-RestMethod -Uri $tmdbUrl -TimeoutSec 5 -ErrorAction Stop
                    
                    if ($tmdbResult -and $tmdbResult.results -and $tmdbResult.results.Count -gt 0) {
                        $movie = $tmdbResult.results[0]
                        
                        # If year was missing, get it from TMDB
                        if (-not $year -and $movie.release_date) {
                            $year = ($movie.release_date -split '-')[0]
                            $fullTitle = "$title ($year)"
                            $displayText = "$Rank. $title`n($year)"
                        }
                        
                        if ($movie.poster_path) {
                            try {
                                $posterUrl = "https://image.tmdb.org/t/p/w92$($movie.poster_path)"
                                $WebClient.Headers.Add("User-Agent", "Mozilla/5.0")
                                $imageBytes = $WebClient.DownloadData($posterUrl)
                                $ms = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
                                $originalImage = [System.Drawing.Image]::FromStream($ms)
                                
                                $resizedImage = New-Object System.Drawing.Bitmap(80, 120)
                                $graphics = [System.Drawing.Graphics]::FromImage($resizedImage)
                                $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                                $graphics.DrawImage($originalImage, 0, 0, 80, 120)
                                $graphics.Dispose()
                                
                                $ImageList.Images.Add($fullTitle, $resizedImage)
                                $originalImage.Dispose()
                                $ms.Dispose()
                                $posterLoaded = $true
                            }
                            catch {
                                # Image download failed
                            }
                        }
                    }
                }
                catch {
                    # TMDB API failed
                }
                
                # Add placeholder if poster didn't load
                if (-not $posterLoaded) {
                    $placeholder = New-Object System.Drawing.Bitmap(80, 120)
                    $g = [System.Drawing.Graphics]::FromImage($placeholder)
                    $g.Clear([System.Drawing.Color]::FromArgb(60, 60, 60))
                    # Draw movie icon on placeholder
                    $font = New-Object System.Drawing.Font("Segoe UI", 24)
                    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(100, 100, 100))
                    $g.DrawString("🎬", $font, $brush, 22, 40)
                    $font.Dispose()
                    $brush.Dispose()
                    $g.Dispose()
                    $ImageList.Images.Add($fullTitle, $placeholder)
                }
                
                $item = New-Object System.Windows.Forms.ListViewItem($displayText)
                $item.ImageKey = $fullTitle
                $MoviesList.Items.Add($item) | Out-Null
                
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $WebClient.Dispose()
            $SourceLabel.Text = "Source: $ListUrl ($Rank movies found)"
        }
        catch {
            $LoadingLabel.Text = "Failed to load: $_"
            $LoadingLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })

$ExportButton.Add_Click({
        [void]$Grid.EndEdit()
    
        $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $SaveDialog.FileName = "DiskSpaceReport.csv"
    
        if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $Global:DataTable | Select-Object Name, SizeMB, SizeGB, Files, Path | Export-Csv -Path $SaveDialog.FileName -NoTypeInformation
                [System.Windows.Forms.MessageBox]::Show("Export successful!", "Success", "OK", "Information") | Out-Null
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Export failed: $_", "Error", "OK", "Error") | Out-Null
            }
        }
    })

$CopyButton.Add_Click({
        [void]$Grid.EndEdit()

        $RowsToCopy = Get-SelectedRows

        if ($RowsToCopy.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No folders selected to copy.", "Warning", "OK", "Warning") | Out-Null
            return
        }

        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select Destination Folder"
        $FolderBrowser.ShowNewFolderButton = $true

        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $DestRoot = $FolderBrowser.SelectedPath

            # Calculate Total Files and Size
            $TotalFilesToProcess = ($RowsToCopy | ForEach-Object { $_.Files } | Measure-Object -Sum).Sum
            $TotalSizeGB = [math]::Round((($RowsToCopy | ForEach-Object { $_.SizeGB } | Measure-Object -Sum).Sum), 2)
            if ($TotalFilesToProcess -eq 0) { $TotalFilesToProcess = 1 }

            # Create Progress Dialog
            $ProgressForm = New-Object System.Windows.Forms.Form
            $ProgressForm.Text = "Copying Files..."
            $ProgressForm.Size = New-Object System.Drawing.Size(500, 220)
            $ProgressForm.StartPosition = "CenterScreen"
            $ProgressForm.FormBorderStyle = "FixedDialog"
            $ProgressForm.MaximizeBox = $false
            $ProgressForm.MinimizeBox = $false
            $ProgressForm.ControlBox = $false
            $ProgressForm.BackColor = $ColorBackground
            $ProgressForm.ForeColor = "White"
            $ProgressForm.TopMost = $true

            $TitleLbl = New-Object System.Windows.Forms.Label
            $TitleLbl.Text = "Copying $($RowsToCopy.Count) folder(s) - $TotalSizeGB GB"
            $TitleLbl.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $TitleLbl.Location = New-Object System.Drawing.Point(20, 15)
            $TitleLbl.AutoSize = $true
            $ProgressForm.Controls.Add($TitleLbl)

            $CurrentFileLbl = New-Object System.Windows.Forms.Label
            $CurrentFileLbl.Text = "Preparing..."
            $CurrentFileLbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $CurrentFileLbl.ForeColor = [System.Drawing.Color]::LightGray
            $CurrentFileLbl.Location = New-Object System.Drawing.Point(20, 50)
            $CurrentFileLbl.Size = New-Object System.Drawing.Size(450, 20)
            $ProgressForm.Controls.Add($CurrentFileLbl)

            $ProgressBarDlg = New-Object System.Windows.Forms.ProgressBar
            $ProgressBarDlg.Location = New-Object System.Drawing.Point(20, 80)
            $ProgressBarDlg.Size = New-Object System.Drawing.Size(445, 25)
            $ProgressBarDlg.Style = "Continuous"
            $ProgressForm.Controls.Add($ProgressBarDlg)

            $ProgressLbl = New-Object System.Windows.Forms.Label
            $ProgressLbl.Text = "0 / $TotalFilesToProcess files (0%)"
            $ProgressLbl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $ProgressLbl.Location = New-Object System.Drawing.Point(20, 115)
            $ProgressLbl.AutoSize = $true
            $ProgressForm.Controls.Add($ProgressLbl)

            $TimeLbl = New-Object System.Windows.Forms.Label
            $TimeLbl.Text = "Estimated time remaining: Calculating..."
            $TimeLbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $TimeLbl.ForeColor = $ColorAccent
            $TimeLbl.Location = New-Object System.Drawing.Point(20, 145)
            $TimeLbl.AutoSize = $true
            $ProgressForm.Controls.Add($TimeLbl)

            # Show the form non-blocking
            $ProgressForm.Show()
            $ProgressForm.Refresh()

            $FolderCount = $RowsToCopy.Count
            $CurrentFolder = 0
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

            foreach ($Row in $RowsToCopy) {
                $CurrentFolder++
                $SourceRoot = $Row.Path
                $FolderName = $Row.Name
                $TargetRoot = Join-Path -Path $DestRoot -ChildPath $FolderName

                $Percent = [int](($CurrentFolder / $FolderCount) * 100)
                $ProgressBarDlg.Value = $Percent
                $CurrentFileLbl.Text = "Copying folder: $FolderName"
                $ProgressLbl.Text = "Folder $CurrentFolder of $FolderCount ($Percent%)"
            
                # Calculate ETR
                $Elapsed = $Stopwatch.Elapsed.TotalSeconds
                if ($CurrentFolder -gt 1 -and $Elapsed -gt 0) {
                    $AvgPerFolder = $Elapsed / ($CurrentFolder - 1)
                    $RemainingFolders = $FolderCount - $CurrentFolder + 1
                    $SecondsLeft = $AvgPerFolder * $RemainingFolders
                    $TimeSpan = [TimeSpan]::FromSeconds($SecondsLeft)
                    $ETR = if ($SecondsLeft -ge 3600) { "{0:hh}:{0:mm}:{0:ss}" -f $TimeSpan } else { "{0:mm}:{0:ss}" -f $TimeSpan }
                    $TimeLbl.Text = "Estimated time remaining: $ETR"
                }
                
                [System.Windows.Forms.Application]::DoEvents()

                # Use robocopy for fast copying
                try {
                    $robocopyArgs = "`"$SourceRoot`" `"$TargetRoot`" /E /R:1 /W:1 /NP /NDL /NJH /NJS /NFL"
                    $process = Start-Process -FilePath "robocopy" -ArgumentList $robocopyArgs -NoNewWindow -PassThru
                    
                    # Poll until complete, updating UI
                    while (-not $process.HasExited) {
                        Start-Sleep -Milliseconds 100
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    
                    if ($process.ExitCode -ge 8) {
                        [System.Windows.Forms.MessageBox]::Show("Warning: Some errors occurred copying $FolderName", "Warning", "OK", "Warning") | Out-Null
                    }
                }
                catch {
                    try {
                        Copy-Item -Path $SourceRoot -Destination $TargetRoot -Recurse -Force -ErrorAction Stop
                    }
                    catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to copy $FolderName : $_", "Error", "OK", "Error") | Out-Null
                    }
                }
                
                [System.Windows.Forms.Application]::DoEvents()
            }

            $Stopwatch.Stop()
            $ProgressForm.Close()
            $ProgressForm.Dispose()

            $ElapsedTime = [TimeSpan]::FromSeconds($Stopwatch.Elapsed.TotalSeconds)
            $TimeStr = if ($ElapsedTime.TotalMinutes -ge 1) { "{0:mm}:{0:ss}" -f $ElapsedTime } else { "$([int]$ElapsedTime.TotalSeconds) seconds" }
        
            [System.Windows.Forms.MessageBox]::Show("Copy complete!`n`nCopied $FolderCount folder(s) to:`n$DestRoot`n`nTime: $TimeStr", "Success", "OK", "Information") | Out-Null
        }
    })

$DeleteButton.Add_Click({
        [void]$Grid.EndEdit()

        $RowsToDelete = Get-SelectedRows

        if ($RowsToDelete.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No folders selected.", "Warning", "OK", "Warning") | Out-Null
            return
        }

        $Count = $RowsToDelete.Count
        $Message = "Are you sure you want to PERMANENTLY DELETE these $Count folders?`n`n" + (($RowsToDelete | Select-Object -First 10 | ForEach-Object { $_.Name }) -join "`n")
        if ($Count -gt 10) { $Message += "`n...and $($Count - 10) more." }

        $Confirm = [System.Windows.Forms.MessageBox]::Show($Message, "Confirm Deletion", "YesNo", "Warning")

        if ($Confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $TotalFilesToProcess = ($RowsToDelete | ForEach-Object { $_.Files } | Measure-Object -Sum).Sum
            if ($TotalFilesToProcess -eq 0) { $TotalFilesToProcess = 1 }

            $StatusLabel.Visible = $true
            $ProgressBar.Visible = $true
            $ProgressBar.Value = 0

            $ProcessedCount = 0
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

            # Collect row indices to delete (in reverse order to avoid index shifting)
            $IndicesToDelete = @()

            foreach ($Row in $RowsToDelete) {
                $FolderPath = $Row.Path
                $RowIndex = $Row.RowIndex

                Get-ChildItem -Path $FolderPath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
                    $File = $_
                    $ProcessedCount++

                    try {
                        Remove-Item -Path $File.FullName -Force -ErrorAction Stop
                    }
                    catch { }

                    if ($ProcessedCount % 5 -eq 0 -or $ProcessedCount -eq $TotalFilesToProcess) {
                        $Elapsed = $Stopwatch.Elapsed.TotalSeconds
                        $Rate = if ($Elapsed -gt 0) { $ProcessedCount / $Elapsed } else { 0 }
                        $RemainingFiles = $TotalFilesToProcess - $ProcessedCount
                        $SecondsLeft = if ($Rate -gt 0) { $RemainingFiles / $Rate } else { 0 }
                        $TimeSpan = [TimeSpan]::FromSeconds($SecondsLeft)
                        $ETR = "{0:mm}:{0:ss}" -f $TimeSpan

                        $ProgressBar.Value = [int](($ProcessedCount / $TotalFilesToProcess) * 100)
                        $StatusLabel.Text = "Deleting $ProcessedCount of $TotalFilesToProcess files... (ETR: $ETR)"
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }

                try {
                    Remove-Item -Path $FolderPath -Recurse -Force -ErrorAction Stop
                    $IndicesToDelete += $RowIndex
                    
                    # Try to remove from Radarr monitoring
                    $FolderName = $Row.Name
                    Remove-MovieFromRadarr -FolderPath $FolderPath -FolderName $FolderName | Out-Null
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Failed to delete $($Row.Name): $_", "Error", "OK", "Error") | Out-Null
                }
            }

            # Delete rows from DataTable in reverse order
            $IndicesToDelete | Sort-Object -Descending | ForEach-Object {
                $Global:DataTable.Rows[$_].Delete()
            }
            $Global:DataTable.AcceptChanges()

            $Stopwatch.Stop()

            $StatusLabel.Text = "Deletion complete."
            $ProgressBar.Visible = $false
            Update-SelectionSummary
            [System.Windows.Forms.MessageBox]::Show("Deletion complete. Matching movies removed from Radarr.", "Success", "OK", "Information") | Out-Null
            
            # Refresh the list to show updated folder list and drive space
            $RefreshButton.PerformClick()
        }
    })

# Select All Button Click
$SelectAllButton.Add_Click({
        foreach ($Row in $Global:DataTable.Rows) {
            $Row["Select"] = $true
        }
        $Grid.Refresh()
        Update-SelectionSummary
    })

# Unselect All Button Click
$UnselectAllButton.Add_Click({
        foreach ($Row in $Global:DataTable.Rows) {
            $Row["Select"] = $false
        }
        $Grid.Refresh()
        Update-SelectionSummary
    })

# Rename Button Click
$RenameButton.Add_Click({
        [void]$Grid.EndEdit()
    
        $SelectedRows = Get-SelectedRows
    
        if ($SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a folder to rename.", "Warning", "OK", "Warning") | Out-Null
            return
        }
    
        # Helper function to perform a single rename
        $DoRename = {
            param($RowData, $NewName)
        
            $CurrentPath = $RowData.Path
            $RowIndex = $RowData.RowIndex
        
            # Validate new name
            $InvalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            foreach ($Char in $InvalidChars) {
                if ($NewName.Contains($Char)) {
                    return "Invalid characters in name: $NewName"
                }
            }
        
            # Calculate new path
            $ParentPath = Split-Path -Path $CurrentPath -Parent
            $NewPath = Join-Path -Path $ParentPath -ChildPath $NewName
        
            # Check if destination exists
            if (Test-Path $NewPath) {
                return "Folder already exists: $NewName"
            }
        
            try {
                Rename-Item -Path $CurrentPath -NewName $NewName -ErrorAction Stop
            
                # Update Grid and DataTable
                $Grid.Rows[$RowIndex].Cells["Name"].Value = $NewName
                $Grid.Rows[$RowIndex].Cells["Path"].Value = $NewPath
                $Global:DataTable.Rows[$RowIndex]["Name"] = $NewName
                $Global:DataTable.Rows[$RowIndex]["Path"] = $NewPath
            
                return $null  # Success
            }
            catch {
                return "Failed: $_"
            }
        }
    
        if ($SelectedRows.Count -eq 1) {
            # Single item - show rename dialog
            $SelectedRow = $SelectedRows[0]
            $CurrentName = $SelectedRow.Name
            $SuggestedName = Extract-MovieName -FolderName $CurrentName
        
            $NewName = Show-RenameDialog -CurrentName $CurrentName -SuggestedName $SuggestedName
        
            if ($NewName -and $NewName -ne $CurrentName) {
                $RenameError = & $DoRename $SelectedRow $NewName
                if ($RenameError) {
                    [System.Windows.Forms.MessageBox]::Show($Error, "Error", "OK", "Error") | Out-Null
                }
                else {
                    $Grid.Refresh()
                    [System.Windows.Forms.MessageBox]::Show("Folder renamed successfully!", "Success", "OK", "Information") | Out-Null
                }
            }
        }
        else {
            # Multiple items - show option dialog
            $OptionForm = New-Object System.Windows.Forms.Form
            $OptionForm.Text = "Rename Multiple Items"
            $OptionForm.Size = New-Object System.Drawing.Size(450, 180)
            $OptionForm.StartPosition = "CenterParent"
            $OptionForm.FormBorderStyle = "FixedDialog"
            $OptionForm.MaximizeBox = $false
            $OptionForm.MinimizeBox = $false
            $OptionForm.BackColor = $ColorBackground
            $OptionForm.ForeColor = "White"
        
            $MsgLabel = New-Object System.Windows.Forms.Label
            $MsgLabel.Text = "You have selected $($SelectedRows.Count) folders to rename.`nHow would you like to proceed?"
            $MsgLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $MsgLabel.Location = New-Object System.Drawing.Point(20, 20)
            $MsgLabel.Size = New-Object System.Drawing.Size(400, 45)
            $OptionForm.Controls.Add($MsgLabel)
        
            $AutoButton = New-Object System.Windows.Forms.Button
            $AutoButton.Text = "Rename All Automatically"
            $AutoButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $AutoButton.BackColor = $ColorSuccess
            $AutoButton.ForeColor = "White"
            $AutoButton.FlatStyle = "Flat"
            $AutoButton.FlatAppearance.BorderSize = 0
            $AutoButton.Size = New-Object System.Drawing.Size(180, 35)
            $AutoButton.Location = New-Object System.Drawing.Point(20, 80)
            $AutoButton.Add_Click({ $OptionForm.Tag = "Auto"; $OptionForm.Close() })
            $OptionForm.Controls.Add($AutoButton)
        
            $IndividualButton = New-Object System.Windows.Forms.Button
            $IndividualButton.Text = "Rename Each Individually"
            $IndividualButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $IndividualButton.BackColor = $ColorAccent
            $IndividualButton.ForeColor = "White"
            $IndividualButton.FlatStyle = "Flat"
            $IndividualButton.FlatAppearance.BorderSize = 0
            $IndividualButton.Size = New-Object System.Drawing.Size(180, 35)
            $IndividualButton.Location = New-Object System.Drawing.Point(210, 80)
            $IndividualButton.Add_Click({ $OptionForm.Tag = "Individual"; $OptionForm.Close() })
            $OptionForm.Controls.Add($IndividualButton)
        
            $CancelBtn = New-Object System.Windows.Forms.Button
            $CancelBtn.Text = "Cancel"
            $CancelBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $CancelBtn.BackColor = $ColorPanel
            $CancelBtn.ForeColor = "White"
            $CancelBtn.FlatStyle = "Flat"
            $CancelBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
            $CancelBtn.Size = New-Object System.Drawing.Size(80, 35)
            $CancelBtn.Location = New-Object System.Drawing.Point(350, 125)
            $CancelBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
            $CancelBtn.Add_Click({ $OptionForm.Tag = "Cancel"; $OptionForm.Close() })
            $OptionForm.Controls.Add($CancelBtn)
        
            $OptionForm.ShowDialog() | Out-Null
            $Choice = $OptionForm.Tag
            $OptionForm.Dispose()
        
            if ($Choice -eq "Cancel" -or -not $Choice) {
                return
            }
        
            $SuccessCount = 0
            $FailCount = 0
            $Errors = @()
        
            if ($Choice -eq "Auto") {
                # Rename all automatically
                foreach ($Row in $SelectedRows) {
                    $CurrentName = $Row.Name
                    $NewName = Extract-MovieName -FolderName $CurrentName
                
                    if ($NewName -ne $CurrentName) {
                        $RenameError = & $DoRename $Row $NewName
                        if ($RenameError) {
                            $FailCount++
                            $Errors += "$CurrentName : $Error"
                        }
                        else {
                            $SuccessCount++
                        }
                    }
                }
            
                $Grid.Refresh()
            
                $ResultMsg = "Renamed $SuccessCount folders successfully."
                if ($FailCount -gt 0) {
                    $ResultMsg += "`n`nFailed to rename $FailCount folders:`n" + ($Errors | Select-Object -First 5 | Out-String)
                }
                [System.Windows.Forms.MessageBox]::Show($ResultMsg, "Rename Complete", "OK", "Information") | Out-Null
            }
            elseif ($Choice -eq "Individual") {
                # Rename each individually with dialog
                foreach ($Row in $SelectedRows) {
                    $CurrentName = $Row.Name
                    $SuggestedName = Extract-MovieName -FolderName $CurrentName
                
                    $NewName = Show-RenameDialog -CurrentName $CurrentName -SuggestedName $SuggestedName
                
                    if (-not $NewName) {
                        # User cancelled this one or clicked Cancel
                        $SkipConfirm = [System.Windows.Forms.MessageBox]::Show("Skip remaining items?", "Continue?", "YesNo", "Question")
                        if ($SkipConfirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                            break
                        }
                        continue
                    }
                
                    if ($NewName -ne $CurrentName) {
                        $RenameError = & $DoRename $Row $NewName
                        if ($RenameError) {
                            $FailCount++
                            [System.Windows.Forms.MessageBox]::Show($Error, "Rename Failed", "OK", "Error") | Out-Null
                        }
                        else {
                            $SuccessCount++
                        }
                    }
                }
            
                $Grid.Refresh()
            
                if ($SuccessCount -gt 0 -or $FailCount -gt 0) {
                    [System.Windows.Forms.MessageBox]::Show("Renamed $SuccessCount folders. Failed: $FailCount", "Complete", "OK", "Information") | Out-Null
                }
            }
        }
    })

# Show Form
$Form.ShowDialog() | Out-Null
$Form.Dispose()
