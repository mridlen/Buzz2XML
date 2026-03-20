<#
.SYNOPSIS
    GUI frontend for Buzz2XML - Jeskola Buzz file converter, VST path remapper, and machine manager.
.DESCRIPTION
    Launches a Windows Forms GUI that wraps the Buzz2XML.ps1 command-line tool.
    Provides tabs for Decode, Encode, Remap, and Machines operations with file browse dialogs.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================================
# Logging
# ============================================================================

$script:LogFile = $null

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $Message"
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logLine
    }
}

# ============================================================================
# Locate the CLI script (Buzz2XML.ps1) -- same directory as this GUI script
# ============================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CliScript = Join-Path $ScriptDir "Buzz2XML.ps1"

# ============================================================================
# Helper: Browse for file  -- BrowseFile
# ============================================================================

function Show-FileDialog {
    param(
        [string]$Title,
        [string]$Filter,
        [bool]$Save = $false
    )
    if ($Save) {
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
    } else {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
    }
    $dlg.Title = $Title
    $dlg.Filter = $Filter
    $dlg.RestoreDirectory = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.FileName
    }
    return $null
}

# ============================================================================
# Helper: Run CLI command and capture output  -- RunCliCommand
# ============================================================================

function Invoke-CliCommand {
    param([string]$Arguments, [System.Windows.Forms.TextBox]$OutputBox)

    $OutputBox.Text = "Running...`r`n"
    $OutputBox.Refresh()

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "powershell.exe"
        $psi.Arguments = "-ExecutionPolicy Bypass -File `"$CliScript`" $Arguments"
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.WorkingDirectory = $ScriptDir

        $process = [System.Diagnostics.Process]::Start($psi)
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        $result = ""
        if ($stdout) { $result += $stdout }
        if ($stderr) { $result += "`r`n$stderr" }
        if ($process.ExitCode -eq 0) {
            $result += "`r`nDone!"
        } else {
            $result += "`r`nProcess exited with code $($process.ExitCode)"
        }

        # Normalize line endings to CRLF for WinForms TextBox display
        $result = $result -replace "`r`n", "`n"
        $result = $result -replace "`r", "`n"
        $result = $result -replace "`n", "`r`n"

        $OutputBox.Text = $result
    } catch {
        $OutputBox.Text = "ERROR: $_"
    }
}

# ============================================================================
# Build the main form
# ============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Buzz2XML - Buzz File Converter & VST Path Remapper"
$form.Size = New-Object System.Drawing.Size(700, 580)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ============================================================================
# Tab control
# ============================================================================

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(665, 525)
$form.Controls.Add($tabControl)

# ============================================================================
# Tab 1: Decode (BMX -> XML)
# ============================================================================

$tabDecode = New-Object System.Windows.Forms.TabPage
$tabDecode.Text = "Decode (BMX -> XML)"
$tabControl.TabPages.Add($tabDecode)

# Input file row
$lblDecIn = New-Object System.Windows.Forms.Label
$lblDecIn.Text = "Input BMX File:"
$lblDecIn.Location = New-Object System.Drawing.Point(15, 20)
$lblDecIn.AutoSize = $true
$tabDecode.Controls.Add($lblDecIn)

$txtDecIn = New-Object System.Windows.Forms.TextBox
$txtDecIn.Location = New-Object System.Drawing.Point(15, 40)
$txtDecIn.Size = New-Object System.Drawing.Size(530, 23)
$tabDecode.Controls.Add($txtDecIn)

$btnDecIn = New-Object System.Windows.Forms.Button
$btnDecIn.Text = "Browse..."
$btnDecIn.Location = New-Object System.Drawing.Point(555, 38)
$btnDecIn.Size = New-Object System.Drawing.Size(85, 27)
$btnDecIn.Add_Click({
    $file = Show-FileDialog -Title "Select Buzz BMX File" -Filter "Buzz Files (*.bmx)|*.bmx|All Files (*.*)|*.*"
    if ($file) {
        $txtDecIn.Text = $file
        # Auto-fill output
        $txtDecOut.Text = [System.IO.Path]::ChangeExtension($file, ".xml")
    }
})
$tabDecode.Controls.Add($btnDecIn)

# Output file row
$lblDecOut = New-Object System.Windows.Forms.Label
$lblDecOut.Text = "Output XML File:"
$lblDecOut.Location = New-Object System.Drawing.Point(15, 75)
$lblDecOut.AutoSize = $true
$tabDecode.Controls.Add($lblDecOut)

$txtDecOut = New-Object System.Windows.Forms.TextBox
$txtDecOut.Location = New-Object System.Drawing.Point(15, 95)
$txtDecOut.Size = New-Object System.Drawing.Size(530, 23)
$tabDecode.Controls.Add($txtDecOut)

$btnDecOut = New-Object System.Windows.Forms.Button
$btnDecOut.Text = "Browse..."
$btnDecOut.Location = New-Object System.Drawing.Point(555, 93)
$btnDecOut.Size = New-Object System.Drawing.Size(85, 27)
$btnDecOut.Add_Click({
    $file = Show-FileDialog -Title "Save XML File As" -Filter "XML Files (*.xml)|*.xml|All Files (*.*)|*.*" -Save $true
    if ($file) { $txtDecOut.Text = $file }
})
$tabDecode.Controls.Add($btnDecOut)

# Decode button
$btnDecode = New-Object System.Windows.Forms.Button
$btnDecode.Text = "Decode BMX to XML"
$btnDecode.Location = New-Object System.Drawing.Point(15, 135)
$btnDecode.Size = New-Object System.Drawing.Size(200, 35)
$btnDecode.BackColor = [System.Drawing.Color]::FromArgb(40, 120, 200)
$btnDecode.ForeColor = [System.Drawing.Color]::White
$btnDecode.FlatStyle = "Flat"
$tabDecode.Controls.Add($btnDecode)

# Output log area
$txtDecLog = New-Object System.Windows.Forms.TextBox
$txtDecLog.Location = New-Object System.Drawing.Point(15, 185)
$txtDecLog.Size = New-Object System.Drawing.Size(625, 295)
$txtDecLog.Multiline = $true
$txtDecLog.ScrollBars = "Both"
$txtDecLog.WordWrap = $false
$txtDecLog.ReadOnly = $true
$txtDecLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtDecLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtDecLog.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 200)
$tabDecode.Controls.Add($txtDecLog)

$btnDecode.Add_Click({
    if (-not $txtDecIn.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select an input BMX file.", "Missing Input", "OK", "Warning")
        return
    }
    if (-not $txtDecOut.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output XML file.", "Missing Output", "OK", "Warning")
        return
    }
    $cliArgs = "-Mode decode -InputFile `"$($txtDecIn.Text)`" -OutputFile `"$($txtDecOut.Text)`""
    Invoke-CliCommand -Arguments $cliArgs -OutputBox $txtDecLog
})

# ============================================================================
# Tab 2: Encode (XML -> BMX)
# ============================================================================

$tabEncode = New-Object System.Windows.Forms.TabPage
$tabEncode.Text = "Encode (XML -> BMX)"
$tabControl.TabPages.Add($tabEncode)

# Input file row
$lblEncIn = New-Object System.Windows.Forms.Label
$lblEncIn.Text = "Input XML File:"
$lblEncIn.Location = New-Object System.Drawing.Point(15, 20)
$lblEncIn.AutoSize = $true
$tabEncode.Controls.Add($lblEncIn)

$txtEncIn = New-Object System.Windows.Forms.TextBox
$txtEncIn.Location = New-Object System.Drawing.Point(15, 40)
$txtEncIn.Size = New-Object System.Drawing.Size(530, 23)
$tabEncode.Controls.Add($txtEncIn)

$btnEncIn = New-Object System.Windows.Forms.Button
$btnEncIn.Text = "Browse..."
$btnEncIn.Location = New-Object System.Drawing.Point(555, 38)
$btnEncIn.Size = New-Object System.Drawing.Size(85, 27)
$btnEncIn.Add_Click({
    $file = Show-FileDialog -Title "Select XML File" -Filter "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
    if ($file) {
        $txtEncIn.Text = $file
        # Auto-fill output
        $txtEncOut.Text = [System.IO.Path]::ChangeExtension($file, ".bmx")
    }
})
$tabEncode.Controls.Add($btnEncIn)

# Output file row
$lblEncOut = New-Object System.Windows.Forms.Label
$lblEncOut.Text = "Output BMX File:"
$lblEncOut.Location = New-Object System.Drawing.Point(15, 75)
$lblEncOut.AutoSize = $true
$tabEncode.Controls.Add($lblEncOut)

$txtEncOut = New-Object System.Windows.Forms.TextBox
$txtEncOut.Location = New-Object System.Drawing.Point(15, 95)
$txtEncOut.Size = New-Object System.Drawing.Size(530, 23)
$tabEncode.Controls.Add($txtEncOut)

$btnEncOut = New-Object System.Windows.Forms.Button
$btnEncOut.Text = "Browse..."
$btnEncOut.Location = New-Object System.Drawing.Point(555, 93)
$btnEncOut.Size = New-Object System.Drawing.Size(85, 27)
$btnEncOut.Add_Click({
    $file = Show-FileDialog -Title "Save BMX File As" -Filter "Buzz Files (*.bmx)|*.bmx|All Files (*.*)|*.*" -Save $true
    if ($file) { $txtEncOut.Text = $file }
})
$tabEncode.Controls.Add($btnEncOut)

# Encode button
$btnEncode = New-Object System.Windows.Forms.Button
$btnEncode.Text = "Encode XML to BMX"
$btnEncode.Location = New-Object System.Drawing.Point(15, 135)
$btnEncode.Size = New-Object System.Drawing.Size(200, 35)
$btnEncode.BackColor = [System.Drawing.Color]::FromArgb(40, 160, 80)
$btnEncode.ForeColor = [System.Drawing.Color]::White
$btnEncode.FlatStyle = "Flat"
$tabEncode.Controls.Add($btnEncode)

# Output log area
$txtEncLog = New-Object System.Windows.Forms.TextBox
$txtEncLog.Location = New-Object System.Drawing.Point(15, 185)
$txtEncLog.Size = New-Object System.Drawing.Size(625, 295)
$txtEncLog.Multiline = $true
$txtEncLog.ScrollBars = "Both"
$txtEncLog.WordWrap = $false
$txtEncLog.ReadOnly = $true
$txtEncLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtEncLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtEncLog.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 200)
$tabEncode.Controls.Add($txtEncLog)

$btnEncode.Add_Click({
    if (-not $txtEncIn.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select an input XML file.", "Missing Input", "OK", "Warning")
        return
    }
    if (-not $txtEncOut.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output BMX file.", "Missing Output", "OK", "Warning")
        return
    }
    $cliArgs = "-Mode encode -InputFile `"$($txtEncIn.Text)`" -OutputFile `"$($txtEncOut.Text)`""
    Invoke-CliCommand -Arguments $cliArgs -OutputBox $txtEncLog
})

# ============================================================================
# Tab 3: Remap VST Paths
# ============================================================================

$tabRemap = New-Object System.Windows.Forms.TabPage
$tabRemap.Text = "Remap VST Paths"
$tabControl.TabPages.Add($tabRemap)

# Input file row
$lblRemIn = New-Object System.Windows.Forms.Label
$lblRemIn.Text = "Input BMX File:"
$lblRemIn.Location = New-Object System.Drawing.Point(15, 20)
$lblRemIn.AutoSize = $true
$tabRemap.Controls.Add($lblRemIn)

$txtRemIn = New-Object System.Windows.Forms.TextBox
$txtRemIn.Location = New-Object System.Drawing.Point(15, 40)
$txtRemIn.Size = New-Object System.Drawing.Size(440, 23)
$tabRemap.Controls.Add($txtRemIn)

$btnRemIn = New-Object System.Windows.Forms.Button
$btnRemIn.Text = "Browse..."
$btnRemIn.Location = New-Object System.Drawing.Point(465, 38)
$btnRemIn.Size = New-Object System.Drawing.Size(85, 27)
$btnRemIn.Add_Click({
    $file = Show-FileDialog -Title "Select Buzz BMX File" -Filter "Buzz Files (*.bmx)|*.bmx|All Files (*.*)|*.*"
    if ($file) {
        $txtRemIn.Text = $file
        # Auto-fill output with _remapped suffix
        $dir = [System.IO.Path]::GetDirectoryName($file)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $ext = [System.IO.Path]::GetExtension($file)
        $txtRemOut.Text = Join-Path $dir "${name}_remapped${ext}"
    }
})
$tabRemap.Controls.Add($btnRemIn)

# List Paths button
$btnListPaths = New-Object System.Windows.Forms.Button
$btnListPaths.Text = "List Paths"
$btnListPaths.Location = New-Object System.Drawing.Point(560, 38)
$btnListPaths.Size = New-Object System.Drawing.Size(85, 27)
$btnListPaths.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 160)
$btnListPaths.ForeColor = [System.Drawing.Color]::White
$btnListPaths.FlatStyle = "Flat"
$btnListPaths.Add_Click({
    if (-not $txtRemIn.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select an input BMX file.", "Missing Input", "OK", "Warning")
        return
    }
    $cliArgs = "-Mode remap -InputFile `"$($txtRemIn.Text)`" -ListPaths"
    Invoke-CliCommand -Arguments $cliArgs -OutputBox $txtRemLog
})
$tabRemap.Controls.Add($btnListPaths)

# Remap From row
$lblRemFrom = New-Object System.Windows.Forms.Label
$lblRemFrom.Text = "Old Path Prefix (remap from):"
$lblRemFrom.Location = New-Object System.Drawing.Point(15, 80)
$lblRemFrom.AutoSize = $true
$tabRemap.Controls.Add($lblRemFrom)

$txtRemFrom = New-Object System.Windows.Forms.TextBox
$txtRemFrom.Location = New-Object System.Drawing.Point(15, 100)
$txtRemFrom.Size = New-Object System.Drawing.Size(625, 23)
$tabRemap.Controls.Add($txtRemFrom)

# Remap To row
$lblRemTo = New-Object System.Windows.Forms.Label
$lblRemTo.Text = "New Path Prefix (remap to):"
$lblRemTo.Location = New-Object System.Drawing.Point(15, 135)
$lblRemTo.AutoSize = $true
$tabRemap.Controls.Add($lblRemTo)

$txtRemTo = New-Object System.Windows.Forms.TextBox
$txtRemTo.Location = New-Object System.Drawing.Point(15, 155)
$txtRemTo.Size = New-Object System.Drawing.Size(625, 23)
$tabRemap.Controls.Add($txtRemTo)

# Output file row
$lblRemOut = New-Object System.Windows.Forms.Label
$lblRemOut.Text = "Output BMX File:"
$lblRemOut.Location = New-Object System.Drawing.Point(15, 190)
$lblRemOut.AutoSize = $true
$tabRemap.Controls.Add($lblRemOut)

$txtRemOut = New-Object System.Windows.Forms.TextBox
$txtRemOut.Location = New-Object System.Drawing.Point(15, 210)
$txtRemOut.Size = New-Object System.Drawing.Size(530, 23)
$tabRemap.Controls.Add($txtRemOut)

$btnRemOut = New-Object System.Windows.Forms.Button
$btnRemOut.Text = "Browse..."
$btnRemOut.Location = New-Object System.Drawing.Point(555, 208)
$btnRemOut.Size = New-Object System.Drawing.Size(85, 27)
$btnRemOut.Add_Click({
    $file = Show-FileDialog -Title "Save Remapped BMX File As" -Filter "Buzz Files (*.bmx)|*.bmx|All Files (*.*)|*.*" -Save $true
    if ($file) { $txtRemOut.Text = $file }
})
$tabRemap.Controls.Add($btnRemOut)

# Remap button
$btnRemap = New-Object System.Windows.Forms.Button
$btnRemap.Text = "Remap Paths"
$btnRemap.Location = New-Object System.Drawing.Point(15, 250)
$btnRemap.Size = New-Object System.Drawing.Size(200, 35)
$btnRemap.BackColor = [System.Drawing.Color]::FromArgb(200, 120, 40)
$btnRemap.ForeColor = [System.Drawing.Color]::White
$btnRemap.FlatStyle = "Flat"
$tabRemap.Controls.Add($btnRemap)

# Output log area
$txtRemLog = New-Object System.Windows.Forms.TextBox
$txtRemLog.Location = New-Object System.Drawing.Point(15, 295)
$txtRemLog.Size = New-Object System.Drawing.Size(625, 185)
$txtRemLog.Multiline = $true
$txtRemLog.ScrollBars = "Both"
$txtRemLog.WordWrap = $false
$txtRemLog.ReadOnly = $true
$txtRemLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtRemLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtRemLog.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 200)
$tabRemap.Controls.Add($txtRemLog)

$btnRemap.Add_Click({
    if (-not $txtRemIn.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select an input BMX file.", "Missing Input", "OK", "Warning")
        return
    }
    if (-not $txtRemOut.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output BMX file.", "Missing Output", "OK", "Warning")
        return
    }
    if (-not $txtRemFrom.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the old path prefix to search for.`n`nTip: Click 'List Paths' first to see what paths exist in the file.", "Missing Remap From", "OK", "Warning")
        return
    }
    if (-not $txtRemTo.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the new path prefix to replace with.", "Missing Remap To", "OK", "Warning")
        return
    }
    $cliArgs = "-Mode remap -InputFile `"$($txtRemIn.Text)`" -OutputFile `"$($txtRemOut.Text)`" -RemapFrom `"$($txtRemFrom.Text)`" -RemapTo `"$($txtRemTo.Text)`""
    Invoke-CliCommand -Arguments $cliArgs -OutputBox $txtRemLog
})

# ============================================================================
# Tab 4: Machines (List / Delete)
# ============================================================================

$tabMach = New-Object System.Windows.Forms.TabPage
$tabMach.Text = "Machines"
$tabControl.TabPages.Add($tabMach)

# Input file row
$lblMachIn = New-Object System.Windows.Forms.Label
$lblMachIn.Text = "Input BMX File:"
$lblMachIn.Location = New-Object System.Drawing.Point(15, 20)
$lblMachIn.AutoSize = $true
$tabMach.Controls.Add($lblMachIn)

$txtMachIn = New-Object System.Windows.Forms.TextBox
$txtMachIn.Location = New-Object System.Drawing.Point(15, 40)
$txtMachIn.Size = New-Object System.Drawing.Size(440, 23)
$tabMach.Controls.Add($txtMachIn)

$btnMachIn = New-Object System.Windows.Forms.Button
$btnMachIn.Text = "Browse..."
$btnMachIn.Location = New-Object System.Drawing.Point(465, 38)
$btnMachIn.Size = New-Object System.Drawing.Size(85, 27)
$btnMachIn.Add_Click({
    $file = Show-FileDialog -Title "Select Buzz BMX File" -Filter "Buzz Files (*.bmx)|*.bmx|All Files (*.*)|*.*"
    if ($file) {
        $txtMachIn.Text = $file
        # Auto-fill output with _cleaned suffix
        $dir = [System.IO.Path]::GetDirectoryName($file)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $ext = [System.IO.Path]::GetExtension($file)
        $txtMachOut.Text = Join-Path $dir "${name}_cleaned${ext}"
    }
})
$tabMach.Controls.Add($btnMachIn)

# List Machines button
$btnListMach = New-Object System.Windows.Forms.Button
$btnListMach.Text = "List Machines"
$btnListMach.Location = New-Object System.Drawing.Point(560, 38)
$btnListMach.Size = New-Object System.Drawing.Size(85, 27)
$btnListMach.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 160)
$btnListMach.ForeColor = [System.Drawing.Color]::White
$btnListMach.FlatStyle = "Flat"
$btnListMach.Add_Click({
    if (-not $txtMachIn.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select an input BMX file.", "Missing Input", "OK", "Warning")
        return
    }
    $cliArgs = "-Mode machines -InputFile `"$($txtMachIn.Text)`" -ListMachines"
    Invoke-CliCommand -Arguments $cliArgs -OutputBox $txtMachLog
})
$tabMach.Controls.Add($btnListMach)

# Wildcard pattern row
$lblMachPattern = New-Object System.Windows.Forms.Label
$lblMachPattern.Text = "Delete Pattern (wildcard, e.g. SVerb*):"
$lblMachPattern.Location = New-Object System.Drawing.Point(15, 80)
$lblMachPattern.AutoSize = $true
$tabMach.Controls.Add($lblMachPattern)

$txtMachPattern = New-Object System.Windows.Forms.TextBox
$txtMachPattern.Location = New-Object System.Drawing.Point(15, 100)
$txtMachPattern.Size = New-Object System.Drawing.Size(625, 23)
$tabMach.Controls.Add($txtMachPattern)

# Exact names (one per line)
$lblMachNames = New-Object System.Windows.Forms.Label
$lblMachNames.Text = "Delete Exact Names (one per line):"
$lblMachNames.Location = New-Object System.Drawing.Point(15, 130)
$lblMachNames.AutoSize = $true
$tabMach.Controls.Add($lblMachNames)

$txtMachNames = New-Object System.Windows.Forms.TextBox
$txtMachNames.Location = New-Object System.Drawing.Point(15, 150)
$txtMachNames.Size = New-Object System.Drawing.Size(625, 60)
$txtMachNames.Multiline = $true
$txtMachNames.ScrollBars = "Vertical"
$txtMachNames.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabMach.Controls.Add($txtMachNames)

# Output file row
$lblMachOut = New-Object System.Windows.Forms.Label
$lblMachOut.Text = "Output BMX File:"
$lblMachOut.Location = New-Object System.Drawing.Point(15, 220)
$lblMachOut.AutoSize = $true
$tabMach.Controls.Add($lblMachOut)

$txtMachOut = New-Object System.Windows.Forms.TextBox
$txtMachOut.Location = New-Object System.Drawing.Point(15, 240)
$txtMachOut.Size = New-Object System.Drawing.Size(530, 23)
$tabMach.Controls.Add($txtMachOut)

$btnMachOut = New-Object System.Windows.Forms.Button
$btnMachOut.Text = "Browse..."
$btnMachOut.Location = New-Object System.Drawing.Point(555, 238)
$btnMachOut.Size = New-Object System.Drawing.Size(85, 27)
$btnMachOut.Add_Click({
    $file = Show-FileDialog -Title "Save Cleaned BMX File As" -Filter "Buzz Files (*.bmx)|*.bmx|All Files (*.*)|*.*" -Save $true
    if ($file) { $txtMachOut.Text = $file }
})
$tabMach.Controls.Add($btnMachOut)

# Delete button
$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Delete Machines"
$btnDelete.Location = New-Object System.Drawing.Point(15, 275)
$btnDelete.Size = New-Object System.Drawing.Size(200, 35)
$btnDelete.BackColor = [System.Drawing.Color]::FromArgb(160, 50, 160)
$btnDelete.ForeColor = [System.Drawing.Color]::White
$btnDelete.FlatStyle = "Flat"
$tabMach.Controls.Add($btnDelete)

# Output log area
$txtMachLog = New-Object System.Windows.Forms.TextBox
$txtMachLog.Location = New-Object System.Drawing.Point(15, 320)
$txtMachLog.Size = New-Object System.Drawing.Size(625, 160)
$txtMachLog.Multiline = $true
$txtMachLog.ScrollBars = "Both"
$txtMachLog.WordWrap = $false
$txtMachLog.ReadOnly = $true
$txtMachLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtMachLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtMachLog.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 200)
$tabMach.Controls.Add($txtMachLog)

$btnDelete.Add_Click({
    if (-not $txtMachIn.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select an input BMX file.", "Missing Input", "OK", "Warning")
        return
    }
    if (-not $txtMachOut.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please specify an output BMX file.", "Missing Output", "OK", "Warning")
        return
    }

    $hasPattern = $txtMachPattern.Text.Trim().Length -gt 0
    $nameLines = @($txtMachNames.Text -split "`r?`n" | Where-Object { $_.Trim().Length -gt 0 })
    $hasNames = $nameLines.Count -gt 0

    if (-not $hasPattern -and -not $hasNames) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a wildcard pattern and/or exact machine names to delete.`n`nTip: Click 'List Machines' first to see what machines exist.", "Nothing to Delete", "OK", "Warning")
        return
    }

    # Build CLI arguments
    $cliArgs = "-Mode machines -InputFile `"$($txtMachIn.Text)`" -OutputFile `"$($txtMachOut.Text)`""
    if ($hasPattern) {
        $cliArgs += " -DeletePattern `"$($txtMachPattern.Text.Trim())`""
    }
    if ($hasNames) {
        # Pass each name as a separate array element using PowerShell array syntax
        $nameArgs = ($nameLines | ForEach-Object { "`"$($_.Trim())`"" }) -join ","
        $cliArgs += " -DeleteNames $nameArgs"
    }

    Invoke-CliCommand -Arguments $cliArgs -OutputBox $txtMachLog
})

# ============================================================================
# Show the form
# ============================================================================

Write-Log "Buzz2XML GUI launched"
[void]$form.ShowDialog()
