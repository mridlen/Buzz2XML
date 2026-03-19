<#
.SYNOPSIS
    Encode and decode Jeskola Buzz (.bmx) files to/from XML format,
    with VST path remapping support.

.DESCRIPTION
    This script can convert a Buzz .bmx binary file to an XML representation,
    convert that XML back to a .bmx binary file, and remap VST plugin paths
    embedded in machine data blobs.

.PARAMETER Mode
    "decode" (bmx -> xml), "encode" (xml -> bmx), or "remap" (rewrite VST paths in bmx).

.PARAMETER InputFile
    Path to the input file (.bmx for decode/remap, .xml for encode).

.PARAMETER OutputFile
    Path to the output file (.xml for decode, .bmx for encode/remap).

.PARAMETER RemapFrom
    (remap mode) The path prefix to search for, e.g. "C:\Program Files (x86)\Jeskola\Buzz\Gear\Vst"

.PARAMETER RemapTo
    (remap mode) The replacement path prefix, e.g. "D:\Audio\VST Plugins"

.PARAMETER ListPaths
    (remap mode) If set, just list all VST/DLL paths found in the file without modifying anything.

.PARAMETER Help
    Show the usage/help page.

.EXAMPLE
    .\Buzz2XML.ps1 -Mode decode -InputFile test_buzz.bmx -OutputFile test_buzz.xml
    .\Buzz2XML.ps1 -Mode encode -InputFile test_buzz.xml -OutputFile test_buzz_rebuilt.bmx
    .\Buzz2XML.ps1 -Mode remap -InputFile test_buzz.bmx -OutputFile remapped.bmx -RemapFrom "C:\Program Files (x86)\Jeskola\Buzz\Gear\Vst" -RemapTo "D:\Audio\VST"
    .\Buzz2XML.ps1 -Mode remap -InputFile test_buzz.bmx -ListPaths
    .\Buzz2XML.ps1 -Help
#>

param(
    [string]$Mode,

    [string]$InputFile,

    [string]$OutputFile,

    [string]$RemapFrom,

    [string]$RemapTo,

    [switch]$ListPaths,

    [Alias("h")]
    [switch]$Help,

    # Catch-all for unrecognized flags like /?, /h, /help, --help
    [Parameter(ValueFromRemainingArguments=$true)]
    $ExtraArgs
)

# ============================================================================
# Help display  -- Show-Help
# ============================================================================

function Show-Help {
    Write-Host ""
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host "  Buzz2XML - Jeskola Buzz (.bmx) File Converter & VST Path Remapper" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  DECODE (BMX -> XML):" -ForegroundColor Green
    Write-Host "    .\Buzz2XML.ps1 -Mode decode -InputFile <file.bmx> -OutputFile <file.xml>"
    Write-Host ""
    Write-Host "    Converts a Buzz .bmx binary song file into a human-readable XML file."
    Write-Host "    All sections are parsed: machines, parameters, patterns, sequences,"
    Write-Host "    connections, wavetable, and more. Machine-specific plugin data and"
    Write-Host "    some sections are stored as base64 for round-trip fidelity."
    Write-Host ""
    Write-Host "  ENCODE (XML -> BMX):" -ForegroundColor Green
    Write-Host "    .\Buzz2XML.ps1 -Mode encode -InputFile <file.xml> -OutputFile <file.bmx>"
    Write-Host ""
    Write-Host "    Converts an XML file (previously created by decode) back into a .bmx"
    Write-Host "    binary file. Produces a byte-for-byte identical copy if the XML has"
    Write-Host "    not been modified."
    Write-Host ""
    Write-Host "  REMAP VST PATHS:" -ForegroundColor Green
    Write-Host "    .\Buzz2XML.ps1 -Mode remap -InputFile <file.bmx> -ListPaths"
    Write-Host "    .\Buzz2XML.ps1 -Mode remap -InputFile <file.bmx> -OutputFile <out.bmx> ``"
    Write-Host "        -RemapFrom <old_path_prefix> -RemapTo <new_path_prefix>"
    Write-Host ""
    Write-Host "    Finds and replaces VST plugin paths embedded inside machine data."
    Write-Host "    Use -ListPaths first to see what paths exist in the file."
    Write-Host "    The remap only changes the path prefix, preserving the filename."
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  -Mode <string>         " -NoNewline -ForegroundColor White
    Write-Host "Required. One of: decode, encode, remap"
    Write-Host "  -InputFile <path>      " -NoNewline -ForegroundColor White
    Write-Host "Required. Path to input file (.bmx or .xml)"
    Write-Host "  -OutputFile <path>     " -NoNewline -ForegroundColor White
    Write-Host "Path to output file (required for decode/encode/remap)"
    Write-Host "  -RemapFrom <string>    " -NoNewline -ForegroundColor White
    Write-Host "Old path prefix to search for (remap mode)"
    Write-Host "  -RemapTo <string>      " -NoNewline -ForegroundColor White
    Write-Host "New path prefix to replace with (remap mode)"
    Write-Host "  -ListPaths             " -NoNewline -ForegroundColor White
    Write-Host "List all embedded file paths without modifying (remap mode)"
    Write-Host "  -Help, -h             " -NoNewline -ForegroundColor White
    Write-Host "Show this help page"
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  # Convert a Buzz song to XML" -ForegroundColor DarkGray
    Write-Host "  .\Buzz2XML.ps1 -Mode decode -InputFile mysong.bmx -OutputFile mysong.xml"
    Write-Host ""
    Write-Host "  # Convert it back" -ForegroundColor DarkGray
    Write-Host "  .\Buzz2XML.ps1 -Mode encode -InputFile mysong.xml -OutputFile mysong_copy.bmx"
    Write-Host ""
    Write-Host "  # See what VST paths are in a song" -ForegroundColor DarkGray
    Write-Host "  .\Buzz2XML.ps1 -Mode remap -InputFile mysong.bmx -ListPaths"
    Write-Host ""
    Write-Host "  # Move VSTs to a new folder" -ForegroundColor DarkGray
    Write-Host "  .\Buzz2XML.ps1 -Mode remap -InputFile mysong.bmx -OutputFile fixed.bmx ``"
    Write-Host "      -RemapFrom `"C:\Program Files (x86)\Jeskola\Buzz\Gear\Vst`" ``"
    Write-Host "      -RemapTo `"D:\Audio\VST Plugins`""
    Write-Host ""
    Write-Host "NOTES:" -ForegroundColor Yellow
    Write-Host "  - A .log file is created alongside the output file for diagnostics."
    Write-Host "  - The GUI version can be launched with: .\Buzz2XML-GUI.ps1"
    Write-Host "  - For help: -Help, -h, or run with no arguments"
    Write-Host ""
}

# ============================================================================
# Check for help flags (including /?, /h, /help, --help via ExtraArgs)
# ============================================================================

$showHelp = $false
if ($Help) { $showHelp = $true }
if (-not $Mode -and -not $ListPaths -and -not $Help) { $showHelp = $true }
# Check if /?, /h, /help were passed as -Mode value (PowerShell treats / as param prefix)
$helpValues = @('/?', '/h', '/help', '-help', '-h', '--help')
if ($Mode -in $helpValues) { $showHelp = $true }
if ($InputFile -in $helpValues) { $showHelp = $true }
if ($ExtraArgs) {
    foreach ($arg in $ExtraArgs) {
        if ($arg -in $helpValues) { $showHelp = $true }
    }
}

if ($showHelp) {
    Show-Help
    exit 0
}

# ============================================================================
# Validate required parameters (since they're no longer Mandatory)
# ============================================================================

if (-not $Mode) {
    Write-Host "ERROR: -Mode is required. Use -Help for usage information." -ForegroundColor Red
    exit 1
}
$validModes = @('decode', 'encode', 'remap')
if ($Mode -notin $validModes) {
    # Could be a help flag that got mangled by the shell (e.g., /help -> C:/Program Files/Git/help)
    if ($Mode -match 'help' -or $Mode -match '^\?$' -or $Mode -match '^[A-Z]:[/\\]$') {
        Show-Help
        exit 0
    }
    Write-Host "ERROR: Invalid mode '$Mode'. Must be one of: $($validModes -join ', ')" -ForegroundColor Red
    Write-Host "Use -Help for usage information." -ForegroundColor Red
    exit 1
}
if (-not $InputFile) {
    Write-Host "ERROR: -InputFile is required. Use -Help for usage information." -ForegroundColor Red
    exit 1
}

# ============================================================================
# Logging setup
# ============================================================================

# $LogFile is set in the main entry point section below
$LogFile = $null

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $Message"
    Add-Content -Path $LogFile -Value $logLine
    Write-Host $logLine
}

# ============================================================================
# XML string sanitization
# ============================================================================

function Sanitize-XmlString {
    # Replace control characters (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F) with &#xNN; escapes
    # XML 1.0 only allows TAB (0x09), LF (0x0A), CR (0x0D) as control chars
    param([string]$Str)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $Str.ToCharArray()) {
        $code = [int]$c
        if ($code -lt 0x20 -and $code -ne 0x09 -and $code -ne 0x0A -and $code -ne 0x0D) {
            $sb.Append("&#x$($code.ToString('X2'));") | Out-Null
        } else {
            $sb.Append($c) | Out-Null
        }
    }
    return $sb.ToString()
}

function Desanitize-XmlString {
    # Reverse the sanitization: convert &#xNN; back to actual control chars
    param([string]$Str)
    return [System.Text.RegularExpressions.Regex]::Replace($Str, '&#x([0-9A-Fa-f]{2});', {
        param($match)
        [char][Convert]::ToInt32($match.Groups[1].Value, 16)
    })
}

# ============================================================================
# Binary reader helpers (used by decode)  -- BinaryReaderHelpers.ps1
# ============================================================================

function Read-AsciizString {
    param([byte[]]$Bytes, [ref]$Pos)
    $start = $Pos.Value
    while ($Bytes[$Pos.Value] -ne 0) { $Pos.Value++ }
    # Use Latin1 (ISO 8859-1) to preserve all byte values 0-255 without corruption
    $str = [System.Text.Encoding]::GetEncoding(28591).GetString($Bytes, $start, $Pos.Value - $start)
    $Pos.Value++  # skip null terminator
    return $str
}

function Read-Byte {
    param([byte[]]$Bytes, [ref]$Pos)
    $val = $Bytes[$Pos.Value]
    $Pos.Value++
    return $val
}

function Read-Word {
    param([byte[]]$Bytes, [ref]$Pos)
    $val = [BitConverter]::ToUInt16($Bytes, $Pos.Value)
    $Pos.Value += 2
    return $val
}

function Read-DWord {
    param([byte[]]$Bytes, [ref]$Pos)
    $val = [BitConverter]::ToUInt32($Bytes, $Pos.Value)
    $Pos.Value += 4
    return $val
}

function Read-Int32 {
    param([byte[]]$Bytes, [ref]$Pos)
    $val = [BitConverter]::ToInt32($Bytes, $Pos.Value)
    $Pos.Value += 4
    return $val
}

function Read-Float {
    param([byte[]]$Bytes, [ref]$Pos)
    $val = [BitConverter]::ToSingle($Bytes, $Pos.Value)
    $Pos.Value += 4
    return $val
}

function Read-ByteArray {
    param([byte[]]$Bytes, [ref]$Pos, [int]$Count)
    $arr = New-Object byte[] $Count
    [Array]::Copy($Bytes, $Pos.Value, $arr, 0, $Count)
    $Pos.Value += $Count
    return ,$arr
}

# ============================================================================
# Parameter size helper
# ============================================================================

function Get-ParamByteSize {
    # param type: 0=note(byte), 1=switch(byte), 2=byte, 3=word
    param([int]$ParamType)
    switch ($ParamType) {
        0 { return 1 }  # note
        1 { return 1 }  # switch
        2 { return 1 }  # byte
        3 { return 2 }  # word
        default { return 1 }
    }
}

# ============================================================================
# Decode: BMX -> XML
# ============================================================================

function ConvertFrom-Buzz {
    param([string]$BmxPath, [string]$XmlPath)

    Write-Log "Starting decode: $BmxPath -> $XmlPath"
    $bytes = [System.IO.File]::ReadAllBytes($BmxPath)
    Write-Log "File size: $($bytes.Length) bytes"

    # --- Header ---
    $magic = [System.Text.Encoding]::ASCII.GetString($bytes, 0, 4)
    if ($magic -ne "Buzz") {
        Write-Log "ERROR: Not a Buzz file (magic=$magic)"
        throw "Not a valid Buzz file."
    }

    $numSections = [BitConverter]::ToUInt32($bytes, 4)
    Write-Log "Number of sections: $numSections"

    # --- Section directory ---
    $sections = @()
    for ($i = 0; $i -lt $numSections; $i++) {
        $dirOffset = 8 + $i * 12
        $secName = [System.Text.Encoding]::ASCII.GetString($bytes, $dirOffset, 4)
        $secOffset = [BitConverter]::ToUInt32($bytes, $dirOffset + 4)
        $secSize = [BitConverter]::ToUInt32($bytes, $dirOffset + 8)
        $sections += [PSCustomObject]@{ Name=$secName; Offset=$secOffset; Size=$secSize }
        Write-Log "  Section: $secName offset=0x$($secOffset.ToString('X8')) size=$secSize"
    }

    # --- Build XML document ---
    $xml = New-Object System.Xml.XmlDocument
    $xmlDecl = $xml.CreateXmlDeclaration("1.0", "UTF-8", $null)
    $xml.AppendChild($xmlDecl) | Out-Null

    $root = $xml.CreateElement("BuzzSong")
    $root.SetAttribute("magic", $magic)
    $xml.AppendChild($root) | Out-Null

    # ---- Pre-parse PARA (needed for MACH and PATT, but don't add to XML yet) ----
    $paraInfo = @()  # array of machine param info from PARA section
    $paraSec = $sections | Where-Object { $_.Name -eq "PARA" }
    if ($paraSec) {
        # Parse PARA into a temp XML to extract info, but don't add to root yet
        $tempXml = New-Object System.Xml.XmlDocument
        $tempRoot = $tempXml.CreateElement("temp")
        $tempXml.AppendChild($tempRoot) | Out-Null
        $paraInfo = Parse-PARA $bytes $paraSec.Offset $paraSec.Size $tempXml $tempRoot  # PARA parsing (Parse-PARA)
    }

    # ---- Pre-parse CONN to compute input counts per machine (needed for PATT) ----
    # PATT stores per-pattern input connection data (amp/pan per row per non-hidden input)
    $machineInputCounts = @{}  # machineIndex -> count of non-hidden inputs
    $connSec = $sections | Where-Object { $_.Name -eq "CONN" }
    if ($connSec -and $paraInfo.Count -gt 0) {
        $cPos = [int]$connSec.Offset
        $cPosRef = [ref]$cPos
        $numConn = Read-Word $bytes $cPosRef
        for ($ci = 0; $ci -lt $numConn; $ci++) {
            $srcIdx = [int](Read-Word $bytes $cPosRef)
            $dstIdx = [int](Read-Word $bytes $cPosRef)
            $null = Read-Word $bytes $cPosRef  # amp
            $null = Read-Word $bytes $cPosRef  # pan

            # Check if source machine is hidden (name starts with 0x01)
            $srcName = $paraInfo[$srcIdx].Name
            $isHidden = ($srcName.Length -gt 0) -and ([byte][char]$srcName[0] -eq 1)

            if (-not $isHidden) {
                if (-not $machineInputCounts.ContainsKey($dstIdx)) {
                    $machineInputCounts[$dstIdx] = 0
                }
                $machineInputCounts[$dstIdx]++
            }
        }
        Write-Log "  Pre-parsed CONN: $numConn connections, machines with inputs: $($machineInputCounts.Count)"
    }

    # ---- Parse each section in original order ----
    foreach ($sec in $sections) {
        Write-Log "Parsing section: $($sec.Name)"
        switch ($sec.Name) {
            "BVER" { Parse-BVER $bytes $sec.Offset $sec.Size $xml $root }       # BVER parsing (Parse-BVER)
            "PARA" { Parse-PARA $bytes $sec.Offset $sec.Size $xml $root | Out-Null }
            "MACH" { Parse-MACH $bytes $sec.Offset $sec.Size $xml $root $paraInfo }  # MACH parsing (Parse-MACH)
            "CONN" { Parse-CONN $bytes $sec.Offset $sec.Size $xml $root }       # CONN parsing (Parse-CONN)
            "CONX" { Parse-CONX $bytes $sec.Offset $sec.Size $xml $root }       # CONX parsing (Parse-CONX)
            "MACX" { Parse-MACX $bytes $sec.Offset $sec.Size $xml $root }       # MACX parsing (Parse-MACX)
            "WAVT" { Parse-WAVT $bytes $sec.Offset $sec.Size $xml $root }       # WAVT parsing (Parse-WAVT)
            "WAVE" { Parse-WAVE $bytes $sec.Offset $sec.Size $xml $root }       # WAVE parsing (Parse-WAVE)
            "PATT" { Parse-PATT $bytes $sec.Offset $sec.Size $xml $root $paraInfo $machineInputCounts }  # PATT parsing (Parse-PATT)
            "PAT2" { Parse-GenericSection $bytes $sec.Offset $sec.Size $xml $root "PAT2" }  # PAT2 (raw)
            "PATX" { Parse-GenericSection $bytes $sec.Offset $sec.Size $xml $root "PATX" }  # PATX (raw)
            "SEQU" { Parse-SEQU $bytes $sec.Offset $sec.Size $xml $root }       # SEQU parsing (Parse-SEQU)
            "BLAH" { Parse-BLAH $bytes $sec.Offset $sec.Size $xml $root }       # BLAH parsing (Parse-BLAH)
            "PDLG" { Parse-PDLG $bytes $sec.Offset $sec.Size $xml $root }       # PDLG parsing (Parse-PDLG)
            "MIDI" { Parse-MIDI $bytes $sec.Offset $sec.Size $xml $root }       # MIDI parsing (Parse-MIDI)
            "BGUI" { Parse-BGUI $bytes $sec.Offset $sec.Size $xml $root }       # BGUI parsing (Parse-BGUI)
            "CWAV" { Parse-WAVE $bytes $sec.Offset $sec.Size $xml $root }       # CWAV parsing (Parse-WAVE)
            default { Parse-GenericSection $bytes $sec.Offset $sec.Size $xml $root $sec.Name }  # unknown section
        }
    }

    # ---- Save XML ----
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.IndentChars = "  "
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.CheckCharacters = $false

    $writer = [System.Xml.XmlWriter]::Create($XmlPath, $settings)
    $xml.Save($writer)
    $writer.Close()

    Write-Log "Decode complete. XML written to $XmlPath"
}

# ============================================================================
# Section parsers (decode)
# ============================================================================

# --- BVER ---
function Parse-BVER {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos
    $version = Read-AsciizString $Bytes $posRef
    $el = $Xml.CreateElement("BVER")
    $el.SetAttribute("version", (Sanitize-XmlString $version))
    $Parent.AppendChild($el) | Out-Null
    Write-Log "  BVER: $version"
}

# --- PARA ---
function Parse-PARA {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $paraEl = $Xml.CreateElement("PARA")
    $Parent.AppendChild($paraEl) | Out-Null

    $numMachines = Read-DWord $Bytes $posRef
    $paraEl.SetAttribute("numMachines", $numMachines)

    $machineInfoList = @()

    for ($m = 0; $m -lt $numMachines; $m++) {
        $machEl = $Xml.CreateElement("Machine")
        $paraEl.AppendChild($machEl) | Out-Null

        $name = Read-AsciizString $Bytes $posRef
        $type = Read-AsciizString $Bytes $posRef
        $numGlobal = Read-DWord $Bytes $posRef
        $numTrack = Read-DWord $Bytes $posRef

        $machEl.SetAttribute("name", (Sanitize-XmlString $name))
        $machEl.SetAttribute("type", (Sanitize-XmlString $type))
        $machEl.SetAttribute("numGlobalParams", $numGlobal)
        $machEl.SetAttribute("numTrackParams", $numTrack)

        Write-Log "  PARA Machine: $name type=$type globals=$numGlobal tracks=$numTrack"

        $globalParams = @()
        $trackParams = @()

        $totalParams = $numGlobal + $numTrack
        for ($p = 0; $p -lt $totalParams; $p++) {
            $paramEl = $Xml.CreateElement("Parameter")
            $machEl.AppendChild($paramEl) | Out-Null

            $ptype = Read-Byte $Bytes $posRef
            $pname = Read-AsciizString $Bytes $posRef
            $minVal = Read-Int32 $Bytes $posRef
            $maxVal = Read-Int32 $Bytes $posRef
            $noVal = Read-Int32 $Bytes $posRef
            $flags = Read-Int32 $Bytes $posRef
            $defVal = Read-Int32 $Bytes $posRef

            $scope = if ($p -lt $numGlobal) { "global" } else { "track" }
            $paramEl.SetAttribute("scope", $scope)
            $paramEl.SetAttribute("type", $ptype)
            $paramEl.SetAttribute("name", (Sanitize-XmlString $pname))
            $paramEl.SetAttribute("minValue", $minVal)
            $paramEl.SetAttribute("maxValue", $maxVal)
            $paramEl.SetAttribute("noValue", $noVal)
            $paramEl.SetAttribute("flags", $flags)
            $paramEl.SetAttribute("defValue", $defVal)

            $paramInfo = [PSCustomObject]@{ Type=$ptype; Name=$pname; ByteSize=(Get-ParamByteSize $ptype) }
            if ($p -lt $numGlobal) {
                $globalParams += $paramInfo
            } else {
                $trackParams += $paramInfo
            }
        }

        $globalSize = ($globalParams | ForEach-Object { $_.ByteSize } | Measure-Object -Sum).Sum
        if ($null -eq $globalSize) { $globalSize = 0 }
        $trackSize = ($trackParams | ForEach-Object { $_.ByteSize } | Measure-Object -Sum).Sum
        if ($null -eq $trackSize) { $trackSize = 0 }

        $machineInfoList += [PSCustomObject]@{
            Name = $name
            GlobalParams = $globalParams
            TrackParams = $trackParams
            GlobalSize = $globalSize
            TrackSize = $trackSize
        }
    }

    return ,$machineInfoList
}

# --- MACH ---
function Parse-MACH {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent, [array]$ParaInfo)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $machEl = $Xml.CreateElement("MACH")
    $Parent.AppendChild($machEl) | Out-Null

    $numMachines = Read-Word $Bytes $posRef
    $machEl.SetAttribute("numMachines", $numMachines)

    for ($m = 0; $m -lt $numMachines; $m++) {
        $mEl = $Xml.CreateElement("Machine")
        $machEl.AppendChild($mEl) | Out-Null

        $name = Read-AsciizString $Bytes $posRef
        $type = Read-Byte $Bytes $posRef
        $mEl.SetAttribute("name", (Sanitize-XmlString $name))
        $mEl.SetAttribute("type", $type)

        $dll = ""
        if ($type -eq 1 -or $type -eq 2) {
            $dll = Read-AsciizString $Bytes $posRef
            $mEl.SetAttribute("dll", (Sanitize-XmlString $dll))
        }

        $x = Read-Float $Bytes $posRef
        $y = Read-Float $Bytes $posRef
        # Store floats as round-trip strings with hex backup for exact precision
        $xBytes = [BitConverter]::GetBytes([float]$x)
        $yBytes = [BitConverter]::GetBytes([float]$y)
        $mEl.SetAttribute("x", $x.ToString("R"))
        $mEl.SetAttribute("y", $y.ToString("R"))
        $mEl.SetAttribute("xHex", [BitConverter]::ToString($xBytes).Replace("-",""))
        $mEl.SetAttribute("yHex", [BitConverter]::ToString($yBytes).Replace("-",""))

        $dataSize = Read-DWord $Bytes $posRef
        $mEl.SetAttribute("dataSize", $dataSize)
        if ($dataSize -gt 0) {
            $data = Read-ByteArray $Bytes $posRef $dataSize
            $mEl.SetAttribute("data", [Convert]::ToBase64String($data))
        }

        $numAttr = Read-Word $Bytes $posRef
        for ($a = 0; $a -lt $numAttr; $a++) {
            $attrEl = $Xml.CreateElement("Attribute")
            $mEl.AppendChild($attrEl) | Out-Null
            $key = Read-AsciizString $Bytes $posRef
            $val = Read-DWord $Bytes $posRef
            $attrEl.SetAttribute("key", (Sanitize-XmlString $key))
            $attrEl.SetAttribute("value", $val)
        }

        # Global parameter state
        if ($m -lt $ParaInfo.Count) {
            $pInfo = $ParaInfo[$m]
            if ($pInfo.GlobalSize -gt 0) {
                $globalState = Read-ByteArray $Bytes $posRef $pInfo.GlobalSize
                $stateEl = $Xml.CreateElement("GlobalState")
                $stateEl.InnerText = [Convert]::ToBase64String($globalState)
                $mEl.AppendChild($stateEl) | Out-Null
            }

            $numTracks = Read-Word $Bytes $posRef
            $mEl.SetAttribute("numTracks", $numTracks)

            for ($t = 0; $t -lt $numTracks; $t++) {
                if ($pInfo.TrackSize -gt 0) {
                    $trackState = Read-ByteArray $Bytes $posRef $pInfo.TrackSize
                    $tsEl = $Xml.CreateElement("TrackState")
                    $tsEl.SetAttribute("track", $t)
                    $tsEl.InnerText = [Convert]::ToBase64String($trackState)
                    $mEl.AppendChild($tsEl) | Out-Null
                }
            }
        } else {
            # No PARA info for this machine - skip state with heuristic
            Write-Log "  WARNING: No PARA info for machine $m ($name), skipping parameter state"
        }

        Write-Log "  MACH[$m]: $name type=$type dll=$dll"
    }
}

# --- CONN ---
function Parse-CONN {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $connEl = $Xml.CreateElement("CONN")
    $Parent.AppendChild($connEl) | Out-Null

    $numConn = Read-Word $Bytes $posRef
    $connEl.SetAttribute("numConnections", $numConn)

    for ($i = 0; $i -lt $numConn; $i++) {
        $cEl = $Xml.CreateElement("Connection")
        $connEl.AppendChild($cEl) | Out-Null
        $cEl.SetAttribute("source", (Read-Word $Bytes $posRef))
        $cEl.SetAttribute("destination", (Read-Word $Bytes $posRef))
        $cEl.SetAttribute("amp", (Read-Word $Bytes $posRef))
        $cEl.SetAttribute("pan", (Read-Word $Bytes $posRef))
    }
}

# --- CONX ---
function Parse-CONX {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    # CONX extends CONN with additional data per connection
    # Format: word numConnections, then per connection: word src, word dst, word amp, word pan, word numExtraParams, then extra data
    # Since the exact format varies, store as raw base64 for round-trip fidelity
    $pos = [int]$Offset
    $data = New-Object byte[] $Size
    [Array]::Copy($Bytes, $pos, $data, 0, $Size)

    $el = $Xml.CreateElement("CONX")
    $el.InnerText = [Convert]::ToBase64String($data)
    $el.SetAttribute("size", $Size)
    $Parent.AppendChild($el) | Out-Null
}

# --- MACX ---
function Parse-MACX {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $macxEl = $Xml.CreateElement("MACX")
    $Parent.AppendChild($macxEl) | Out-Null

    $numMach = Read-Word $Bytes $posRef
    $macxEl.SetAttribute("numMachines", $numMach)

    for ($m = 0; $m -lt $numMach; $m++) {
        $mEl = $Xml.CreateElement("Machine")
        $macxEl.AppendChild($mEl) | Out-Null

        $name = Read-AsciizString $Bytes $posRef
        $mEl.SetAttribute("name", (Sanitize-XmlString $name))

        $numAttrs = Read-DWord $Bytes $posRef

        for ($a = 0; $a -lt $numAttrs; $a++) {
            $attrEl = $Xml.CreateElement("Attribute")
            $mEl.AppendChild($attrEl) | Out-Null

            $key = Read-AsciizString $Bytes $posRef
            $attrSize = Read-DWord $Bytes $posRef
            $attrData = Read-ByteArray $Bytes $posRef $attrSize

            $attrEl.SetAttribute("key", (Sanitize-XmlString $key))
            $attrEl.SetAttribute("size", $attrSize)
            $attrEl.InnerText = [Convert]::ToBase64String($attrData)
        }
    }
}

# --- WAVT ---
function Parse-WAVT {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $wavtEl = $Xml.CreateElement("WAVT")
    $Parent.AppendChild($wavtEl) | Out-Null

    $numWaves = Read-Word $Bytes $posRef
    $wavtEl.SetAttribute("numWaves", $numWaves)

    for ($w = 0; $w -lt $numWaves; $w++) {
        $wEl = $Xml.CreateElement("Wave")
        $wavtEl.AppendChild($wEl) | Out-Null

        $index = Read-Word $Bytes $posRef
        $fileName = Read-AsciizString $Bytes $posRef
        $waveName = Read-AsciizString $Bytes $posRef
        $volume = Read-Float $Bytes $posRef
        $flags = Read-Byte $Bytes $posRef

        $wEl.SetAttribute("index", $index)
        $wEl.SetAttribute("fileName", (Sanitize-XmlString $fileName))
        $wEl.SetAttribute("name", (Sanitize-XmlString $waveName))
        $wEl.SetAttribute("volume", $volume.ToString("R"))
        $volBytes = [BitConverter]::GetBytes([float]$volume)
        $wEl.SetAttribute("volumeHex", [BitConverter]::ToString($volBytes).Replace("-",""))
        $wEl.SetAttribute("flags", $flags)

        # Envelope data if bit 7 set
        if ($flags -band 0x80) {
            $numEnvelopes = Read-Word $Bytes $posRef
            $wEl.SetAttribute("numEnvelopes", $numEnvelopes)

            for ($e = 0; $e -lt $numEnvelopes; $e++) {
                $envEl = $Xml.CreateElement("Envelope")
                $wEl.AppendChild($envEl) | Out-Null

                $envEl.SetAttribute("attackTime", (Read-Word $Bytes $posRef))
                $envEl.SetAttribute("decayTime", (Read-Word $Bytes $posRef))
                $envEl.SetAttribute("sustainLevel", (Read-Word $Bytes $posRef))
                $envEl.SetAttribute("releaseTime", (Read-Word $Bytes $posRef))
                $envEl.SetAttribute("adsrSubdivide", (Read-Byte $Bytes $posRef))
                $envEl.SetAttribute("adsrFlags", (Read-Byte $Bytes $posRef))

                $numPoints = Read-Word $Bytes $posRef
                $envDisabled = ($numPoints -band 0x8000) -ne 0
                $actualPoints = $numPoints -band 0x7FFF
                $envEl.SetAttribute("numPoints", $actualPoints)
                $envEl.SetAttribute("disabled", $envDisabled.ToString().ToLower())

                for ($p = 0; $p -lt $actualPoints; $p++) {
                    $ptEl = $Xml.CreateElement("Point")
                    $envEl.AppendChild($ptEl) | Out-Null
                    $ptEl.SetAttribute("x", (Read-Word $Bytes $posRef))
                    $ptEl.SetAttribute("y", (Read-Word $Bytes $posRef))
                    $ptEl.SetAttribute("flags", (Read-Byte $Bytes $posRef))
                }
            }
        }

        # Levels
        $numLevels = Read-Byte $Bytes $posRef
        $wEl.SetAttribute("numLevels", $numLevels)

        for ($l = 0; $l -lt $numLevels; $l++) {
            $lvlEl = $Xml.CreateElement("Level")
            $wEl.AppendChild($lvlEl) | Out-Null
            $lvlEl.SetAttribute("numSamples", (Read-DWord $Bytes $posRef))
            $lvlEl.SetAttribute("loopBegin", (Read-DWord $Bytes $posRef))
            $lvlEl.SetAttribute("loopEnd", (Read-DWord $Bytes $posRef))
            $lvlEl.SetAttribute("samplesPerSecond", (Read-DWord $Bytes $posRef))
            $lvlEl.SetAttribute("rootNote", (Read-Byte $Bytes $posRef))
        }
    }
}

# --- WAVE / CWAV ---
function Parse-WAVE {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos
    $endPos = $Offset + $Size

    # Detect section name from the directory
    $secName = [System.Text.Encoding]::ASCII.GetString($Bytes, ($Offset - $Size), 4)
    # Use generic element name
    $waveEl = $Xml.CreateElement("WAVE")
    $Parent.AppendChild($waveEl) | Out-Null

    $numWaves = Read-Word $Bytes $posRef
    $waveEl.SetAttribute("numWaves", $numWaves)

    for ($w = 0; $w -lt $numWaves; $w++) {
        $wEl = $Xml.CreateElement("WaveData")
        $waveEl.AppendChild($wEl) | Out-Null

        $index = Read-Word $Bytes $posRef
        $format = Read-Byte $Bytes $posRef
        $wEl.SetAttribute("index", $index)
        $wEl.SetAttribute("format", $format)

        $numBytes = Read-DWord $Bytes $posRef
        $wEl.SetAttribute("dataSize", $numBytes)

        if ($numBytes -gt 0) {
            $waveData = Read-ByteArray $Bytes $posRef $numBytes
            $wEl.InnerText = [Convert]::ToBase64String($waveData)
        }
    }
}

# --- PATT ---
function Parse-PATT {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent, [array]$ParaInfo, [hashtable]$MachineInputCounts)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $pattEl = $Xml.CreateElement("PATT")
    $Parent.AppendChild($pattEl) | Out-Null

    $numMachines = $ParaInfo.Count
    $endPos = [int]$Offset + [int]$Size

    for ($m = 0; $m -lt $numMachines; $m++) {
        # Bounds check before reading machine header
        if ($pos -ge $endPos) { break }

        $machPattEl = $Xml.CreateElement("MachinePatterns")
        $pattEl.AppendChild($machPattEl) | Out-Null
        $machPattEl.SetAttribute("machineIndex", $m)

        $numPatterns = Read-Word $Bytes $posRef
        $numTracks = Read-Word $Bytes $posRef
        $machPattEl.SetAttribute("numPatterns", $numPatterns)
        $machPattEl.SetAttribute("numTracks", $numTracks)

        $pInfo = $ParaInfo[$m]
        $paramRowSize = $pInfo.GlobalSize + ($pInfo.TrackSize * $numTracks)

        # Get input count for this machine (non-hidden inputs from CONN)
        $inputCount = 0
        if ($MachineInputCounts -and $MachineInputCounts.ContainsKey($m)) {
            $inputCount = $MachineInputCounts[$m]
        }
        # Input data per pattern: inputCount * (2 bytes sourceMachineIndex + rows * 4 bytes amp/pan)
        $machPattEl.SetAttribute("inputCount", $inputCount)

        if ($numPatterns -gt 0) {
            Write-Log "  PATT machine $m ($($pInfo.Name)): $numPatterns patterns, $numTracks tracks, paramRowSize=$paramRowSize, inputCount=$inputCount"
        }

        for ($p = 0; $p -lt $numPatterns; $p++) {
            $patEl = $Xml.CreateElement("Pattern")
            $machPattEl.AppendChild($patEl) | Out-Null

            $patName = Read-AsciizString $Bytes $posRef
            $patLength = Read-Word $Bytes $posRef
            $patEl.SetAttribute("name", (Sanitize-XmlString $patName))
            $patEl.SetAttribute("rows", $patLength)

            # Input connection data: for each non-hidden input, word sourceMachineIndex + rows * (word amp + word pan)
            $inputDataSize = $inputCount * (2 + $patLength * 4)

            # Parameter data: rows * (globalSize + trackSize * numTracks)
            $paramDataSize = $paramRowSize * $patLength

            # Total pattern data = input data + parameter data
            $totalDataSize = $inputDataSize + $paramDataSize
            $patEl.SetAttribute("inputDataSize", $inputDataSize)
            $patEl.SetAttribute("paramDataSize", $paramDataSize)

            # Store all pattern data (input + params) as base64 blob
            if ($totalDataSize -gt 0) {
                if (($pos + $totalDataSize) -gt $Bytes.Length) {
                    Write-Log "WARNING: PATT machine $m pattern $p would read past file end (pos=$pos, need=$totalDataSize, fileLen=$($Bytes.Length))"
                    break
                }
                $patData = Read-ByteArray $Bytes $posRef $totalDataSize
                $patEl.InnerText = [Convert]::ToBase64String($patData)
            }
        }
    }
}

# --- SEQU ---
function Parse-SEQU {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $sequEl = $Xml.CreateElement("SEQU")
    $Parent.AppendChild($sequEl) | Out-Null

    $sequEl.SetAttribute("endOfSong", (Read-DWord $Bytes $posRef))
    $sequEl.SetAttribute("beginLoop", (Read-DWord $Bytes $posRef))
    $sequEl.SetAttribute("endLoop", (Read-DWord $Bytes $posRef))

    $numSeq = Read-Word $Bytes $posRef
    $sequEl.SetAttribute("numSequences", $numSeq)

    for ($s = 0; $s -lt $numSeq; $s++) {
        $sEl = $Xml.CreateElement("Sequence")
        $sequEl.AppendChild($sEl) | Out-Null

        $machIdx = Read-Word $Bytes $posRef
        $numEvents = Read-DWord $Bytes $posRef
        $bytesPerPos = Read-Byte $Bytes $posRef
        $bytesPerEvent = Read-Byte $Bytes $posRef

        $sEl.SetAttribute("machineIndex", $machIdx)
        $sEl.SetAttribute("numEvents", $numEvents)
        $sEl.SetAttribute("bytesPerPos", $bytesPerPos)
        $sEl.SetAttribute("bytesPerEvent", $bytesPerEvent)

        for ($e = 0; $e -lt $numEvents; $e++) {
            $evEl = $Xml.CreateElement("Event")
            $sEl.AppendChild($evEl) | Out-Null

            # Read position (variable size) - cast to [int] to prevent byte overflow on shift
            $eventPos = 0
            for ($b = 0; $b -lt $bytesPerPos; $b++) {
                $eventPos = $eventPos -bor ([int](Read-Byte $Bytes $posRef) -shl ($b * 8))
            }
            # Read event value (variable size) - cast to [int] to prevent byte overflow on shift
            $eventVal = 0
            for ($b = 0; $b -lt $bytesPerEvent; $b++) {
                $eventVal = $eventVal -bor ([int](Read-Byte $Bytes $posRef) -shl ($b * 8))
            }

            $evEl.SetAttribute("pos", $eventPos)
            $evEl.SetAttribute("value", "0x$($eventVal.ToString('X2'))")
        }
    }
}

# --- BLAH ---
function Parse-BLAH {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $el = $Xml.CreateElement("BLAH")
    $Parent.AppendChild($el) | Out-Null

    $numChars = Read-DWord $Bytes $posRef
    if ($numChars -gt 0) {
        $text = [System.Text.Encoding]::ASCII.GetString($Bytes, $pos, $numChars)
        $el.InnerText = Sanitize-XmlString $text
    }
    $el.SetAttribute("length", $numChars)
}

# --- PDLG ---
function Parse-PDLG {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    # Store raw data for round-trip fidelity
    $data = New-Object byte[] $Size
    [Array]::Copy($Bytes, $Offset, $data, 0, $Size)
    $el = $Xml.CreateElement("PDLG")
    $el.InnerText = [Convert]::ToBase64String($data)
    $el.SetAttribute("size", $Size)
    $Parent.AppendChild($el) | Out-Null
}

# --- MIDI ---
function Parse-MIDI {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    $pos = [int]$Offset
    $posRef = [ref]$pos
    $endPos = $Offset + $Size

    $midiEl = $Xml.CreateElement("MIDI")
    $Parent.AppendChild($midiEl) | Out-Null

    # MIDI bindings: list terminated by zero byte
    while ($pos -lt $endPos) {
        if ($Bytes[$pos] -eq 0) {
            # Terminating zero - no bindings
            $pos++
            break
        }

        $bindEl = $Xml.CreateElement("Binding")
        $midiEl.AppendChild($bindEl) | Out-Null

        $machName = Read-AsciizString $Bytes $posRef
        $paramGroup = Read-Byte $Bytes $posRef
        $paramTrack = Read-Byte $Bytes $posRef
        $paramNumber = Read-Byte $Bytes $posRef
        $midiChannel = Read-Byte $Bytes $posRef
        $midiController = Read-Byte $Bytes $posRef

        $bindEl.SetAttribute("machine", (Sanitize-XmlString $machName))
        $bindEl.SetAttribute("paramGroup", $paramGroup)
        $bindEl.SetAttribute("paramTrack", $paramTrack)
        $bindEl.SetAttribute("paramNumber", $paramNumber)
        $bindEl.SetAttribute("midiChannel", $midiChannel)
        $bindEl.SetAttribute("midiController", $midiController)
    }
}

# --- BGUI ---
function Parse-BGUI {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent)
    # Store as raw base64 for round-trip fidelity
    $data = New-Object byte[] $Size
    [Array]::Copy($Bytes, $Offset, $data, 0, $Size)
    $el = $Xml.CreateElement("BGUI")
    $el.InnerText = [Convert]::ToBase64String($data)
    $el.SetAttribute("size", $Size)
    $Parent.AppendChild($el) | Out-Null
}

# --- Generic/Unknown section ---
function Parse-GenericSection {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent, [string]$SectionName)
    $data = New-Object byte[] $Size
    [Array]::Copy($Bytes, $Offset, $data, 0, $Size)
    $el = $Xml.CreateElement($SectionName)
    $el.InnerText = [Convert]::ToBase64String($data)
    $el.SetAttribute("size", $Size)
    $el.SetAttribute("encoding", "base64")
    $Parent.AppendChild($el) | Out-Null
    Write-Log "  ${SectionName}: stored as base64 (${Size} bytes)"
}

# ============================================================================
# Encode: XML -> BMX
# ============================================================================

function ConvertTo-Buzz {
    param([string]$XmlPath, [string]$BmxPath)

    Write-Log "Starting encode: $XmlPath -> $BmxPath"

    $xml = New-Object System.Xml.XmlDocument
    $xml.Load($XmlPath)
    $root = $xml.DocumentElement

    # --- Collect section data ---
    # We need to build each section's binary data, then assemble the file.
    # The section order in the XML determines section order in the file.

    $sectionOrder = @()
    $sectionData = @{}

    foreach ($child in $root.ChildNodes) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $secName = $child.Name
        if ($secName.Length -gt 4) { continue }  # skip non-section elements

        Write-Log "  Encoding section: $secName"
        $data = $null

        switch ($secName) {
            "BVER" { $data = Encode-BVER $child }           # BVER encoding (Encode-BVER)
            "PARA" { $data = Encode-PARA $child }           # PARA encoding (Encode-PARA)
            "MACH" { $data = Encode-MACH $child $root }     # MACH encoding (Encode-MACH)
            "CONN" { $data = Encode-CONN $child }           # CONN encoding (Encode-CONN)
            "SEQU" { $data = Encode-SEQU $child }           # SEQU encoding (Encode-SEQU)
            "WAVT" { $data = Encode-WAVT $child }           # WAVT encoding (Encode-WAVT)
            "WAVE" { $data = Encode-WAVE $child }           # WAVE encoding (Encode-WAVE)
            "PATT" { $data = Encode-PATT $child }     # PATT encoding (Encode-PATT)
            "BLAH" { $data = Encode-BLAH $child }           # BLAH encoding (Encode-BLAH)
            "MIDI" { $data = Encode-MIDI $child }           # MIDI encoding (Encode-MIDI)
            "MACX" { $data = Encode-MACX $child }           # MACX encoding (Encode-MACX)
            default {
                # Base64-encoded raw section (CONX, PDLG, BGUI, PAT2, PATX, etc.)
                $data = Encode-RawSection $child            # raw section encoding (Encode-RawSection)
            }
        }

        if ($null -ne $data) {
            $sectionOrder += $secName
            $sectionData[$secName] = $data
        }
    }

    # --- Assemble file ---
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    # Magic
    $bw.Write([System.Text.Encoding]::ASCII.GetBytes("Buzz"))
    # Number of sections
    $bw.Write([uint32]$sectionOrder.Count)
    # Section directory placeholder (12 bytes * 31 slots)
    $dirStart = $ms.Position
    for ($i = 0; $i -lt 31; $i++) {
        $bw.Write([uint32]0)  # name placeholder
        $bw.Write([uint32]0)  # offset placeholder
        $bw.Write([uint32]0)  # size placeholder
    }

    # Write sections and record offsets
    $sectionOffsets = @{}
    for ($i = 0; $i -lt $sectionOrder.Count; $i++) {
        $secName = $sectionOrder[$i]
        $sectionOffsets[$secName] = [uint32]$ms.Position
        $bw.Write($sectionData[$secName])
    }

    # Go back and fill in section directory
    $ms.Position = $dirStart
    for ($i = 0; $i -lt $sectionOrder.Count; $i++) {
        $secName = $sectionOrder[$i]
        $nameBytes = [System.Text.Encoding]::ASCII.GetBytes($secName)
        # Pad name to 4 bytes
        $padded = New-Object byte[] 4
        [Array]::Copy($nameBytes, $padded, [Math]::Min($nameBytes.Length, 4))
        $bw.Write($padded)
        $bw.Write([uint32]$sectionOffsets[$secName])
        $bw.Write([uint32]$sectionData[$secName].Length)
    }

    # Write to file
    $fileBytes = $ms.ToArray()
    $bw.Close()
    $ms.Close()

    [System.IO.File]::WriteAllBytes($BmxPath, $fileBytes)
    Write-Log "Encode complete. BMX written to $BmxPath ($($fileBytes.Length) bytes)"
}

# ============================================================================
# Section encoders (encode)
# ============================================================================

# Helper: convert hex string to byte array (compatible with older .NET)
function ConvertFrom-HexString {
    param([string]$Hex)
    $bytes = New-Object byte[] ($Hex.Length / 2)
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        $bytes[$i] = [Convert]::ToByte($Hex.Substring($i * 2, 2), 16)
    }
    return ,$bytes
}

# Helper: write asciiz string to MemoryStream (auto-desanitizes XML escapes)
function Write-Asciiz {
    param([System.IO.BinaryWriter]$BW, [string]$Str)
    $raw = Desanitize-XmlString $Str
    $bw.Write([System.Text.Encoding]::GetEncoding(28591).GetBytes($raw))
    $bw.Write([byte]0)
}

# --- BVER ---
function Encode-BVER {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)
    Write-Asciiz $bw $El.GetAttribute("version")
    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- PARA ---
function Encode-PARA {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $machines = @($El.SelectNodes("Machine"))
    $bw.Write([uint32]$machines.Count)

    foreach ($mach in $machines) {
        Write-Asciiz $bw $mach.GetAttribute("name")
        Write-Asciiz $bw $mach.GetAttribute("type")
        $bw.Write([uint32]$mach.GetAttribute("numGlobalParams"))
        $bw.Write([uint32]$mach.GetAttribute("numTrackParams"))

        $params = @($mach.SelectNodes("Parameter"))
        foreach ($param in $params) {
            $bw.Write([byte]$param.GetAttribute("type"))
            Write-Asciiz $bw $param.GetAttribute("name")
            $bw.Write([int32]$param.GetAttribute("minValue"))
            $bw.Write([int32]$param.GetAttribute("maxValue"))
            $bw.Write([int32]$param.GetAttribute("noValue"))
            $bw.Write([int32]$param.GetAttribute("flags"))
            $bw.Write([int32]$param.GetAttribute("defValue"))
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- MACH ---
function Encode-MACH {
    param([System.Xml.XmlElement]$El, [System.Xml.XmlElement]$Root)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $machines = @($El.SelectNodes("Machine"))
    $bw.Write([uint16]$machines.Count)

    foreach ($mach in $machines) {
        Write-Asciiz $bw $mach.GetAttribute("name")
        $type = [byte]$mach.GetAttribute("type")
        $bw.Write($type)

        if ($type -eq 1 -or $type -eq 2) {
            Write-Asciiz $bw $mach.GetAttribute("dll")
        }

        # Use hex values for exact float precision if available
        $xHex = $mach.GetAttribute("xHex")
        $yHex = $mach.GetAttribute("yHex")
        if ($xHex) {
            $xBytes = ConvertFrom-HexString $xHex; $bw.Write($xBytes)
        } else {
            $bw.Write([float]$mach.GetAttribute("x"))
        }
        if ($yHex) {
            $yBytes = ConvertFrom-HexString $yHex; $bw.Write($yBytes)
        } else {
            $bw.Write([float]$mach.GetAttribute("y"))
        }

        $dataSize = [uint32]$mach.GetAttribute("dataSize")
        $bw.Write($dataSize)
        if ($dataSize -gt 0) {
            $data = [Convert]::FromBase64String($mach.GetAttribute("data"))
            $bw.Write($data)
        }

        $attrs = @($mach.SelectNodes("Attribute"))
        $bw.Write([uint16]$attrs.Count)
        foreach ($attr in $attrs) {
            Write-Asciiz $bw $attr.GetAttribute("key")
            $bw.Write([uint32]$attr.GetAttribute("value"))
        }

        # Global state
        $globalStateEl = $mach.SelectSingleNode("GlobalState")
        if ($globalStateEl) {
            $stateBytes = [Convert]::FromBase64String($globalStateEl.InnerText)
            $bw.Write($stateBytes)
        }

        # Num tracks
        $numTracks = 0
        if ($mach.HasAttribute("numTracks")) {
            $numTracks = [uint16]$mach.GetAttribute("numTracks")
        }
        $bw.Write([uint16]$numTracks)

        # Track state
        $trackStates = @($mach.SelectNodes("TrackState"))
        foreach ($ts in $trackStates) {
            $tsBytes = [Convert]::FromBase64String($ts.InnerText)
            $bw.Write($tsBytes)
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- CONN ---
function Encode-CONN {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $conns = @($El.SelectNodes("Connection"))
    $bw.Write([uint16]$conns.Count)

    foreach ($conn in $conns) {
        $bw.Write([uint16]$conn.GetAttribute("source"))
        $bw.Write([uint16]$conn.GetAttribute("destination"))
        $bw.Write([uint16]$conn.GetAttribute("amp"))
        $bw.Write([uint16]$conn.GetAttribute("pan"))
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- SEQU ---
function Encode-SEQU {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $bw.Write([uint32]$El.GetAttribute("endOfSong"))
    $bw.Write([uint32]$El.GetAttribute("beginLoop"))
    $bw.Write([uint32]$El.GetAttribute("endLoop"))

    $seqs = @($El.SelectNodes("Sequence"))
    $bw.Write([uint16]$seqs.Count)

    foreach ($seq in $seqs) {
        $bw.Write([uint16]$seq.GetAttribute("machineIndex"))
        $bw.Write([uint32]$seq.GetAttribute("numEvents"))
        $bytesPerPos = [byte]$seq.GetAttribute("bytesPerPos")
        $bytesPerEvent = [byte]$seq.GetAttribute("bytesPerEvent")
        $bw.Write($bytesPerPos)
        $bw.Write($bytesPerEvent)

        $events = @($seq.SelectNodes("Event"))
        foreach ($ev in $events) {
            $evPos = [int]$ev.GetAttribute("pos")
            for ($b = 0; $b -lt $bytesPerPos; $b++) {
                $bw.Write([byte](($evPos -shr ($b * 8)) -band 0xFF))
            }
            # Parse hex value
            $valStr = $ev.GetAttribute("value")
            $evVal = 0
            if ($valStr.StartsWith("0x")) {
                $evVal = [Convert]::ToInt32($valStr, 16)
            } else {
                $evVal = [int]$valStr
            }
            for ($b = 0; $b -lt $bytesPerEvent; $b++) {
                $bw.Write([byte](($evVal -shr ($b * 8)) -band 0xFF))
            }
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- WAVT ---
function Encode-WAVT {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $waves = @($El.SelectNodes("Wave"))
    $bw.Write([uint16]$waves.Count)

    foreach ($wave in $waves) {
        $bw.Write([uint16]$wave.GetAttribute("index"))
        Write-Asciiz $bw $wave.GetAttribute("fileName")
        Write-Asciiz $bw $wave.GetAttribute("name")
        $volHex = $wave.GetAttribute("volumeHex")
        if ($volHex) {
            $volBytes = ConvertFrom-HexString $volHex; $bw.Write($volBytes)
        } else {
            $bw.Write([float]$wave.GetAttribute("volume"))
        }
        $flags = [byte]$wave.GetAttribute("flags")
        $bw.Write($flags)

        # Envelopes if bit 7 set
        if ($flags -band 0x80) {
            $envelopes = @($wave.SelectNodes("Envelope"))
            $bw.Write([uint16]$envelopes.Count)

            foreach ($env in $envelopes) {
                $bw.Write([uint16]$env.GetAttribute("attackTime"))
                $bw.Write([uint16]$env.GetAttribute("decayTime"))
                $bw.Write([uint16]$env.GetAttribute("sustainLevel"))
                $bw.Write([uint16]$env.GetAttribute("releaseTime"))
                $bw.Write([byte]$env.GetAttribute("adsrSubdivide"))
                $bw.Write([byte]$env.GetAttribute("adsrFlags"))

                $numPoints = [uint16]$env.GetAttribute("numPoints")
                $disabled = $env.GetAttribute("disabled") -eq "true"
                if ($disabled) { $numPoints = $numPoints -bor 0x8000 }
                $bw.Write([uint16]$numPoints)

                $points = @($env.SelectNodes("Point"))
                foreach ($pt in $points) {
                    $bw.Write([uint16]$pt.GetAttribute("x"))
                    $bw.Write([uint16]$pt.GetAttribute("y"))
                    $bw.Write([byte]$pt.GetAttribute("flags"))
                }
            }
        }

        # Levels
        $levels = @($wave.SelectNodes("Level"))
        $bw.Write([byte]$levels.Count)

        foreach ($lvl in $levels) {
            $bw.Write([uint32]$lvl.GetAttribute("numSamples"))
            $bw.Write([uint32]$lvl.GetAttribute("loopBegin"))
            $bw.Write([uint32]$lvl.GetAttribute("loopEnd"))
            $bw.Write([uint32]$lvl.GetAttribute("samplesPerSecond"))
            $bw.Write([byte]$lvl.GetAttribute("rootNote"))
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- WAVE ---
function Encode-WAVE {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $waveDatas = @($El.SelectNodes("WaveData"))
    $bw.Write([uint16]$waveDatas.Count)

    foreach ($wd in $waveDatas) {
        $bw.Write([uint16]$wd.GetAttribute("index"))
        $bw.Write([byte]$wd.GetAttribute("format"))
        $dataSize = [uint32]$wd.GetAttribute("dataSize")
        $bw.Write($dataSize)
        if ($dataSize -gt 0) {
            $data = [Convert]::FromBase64String($wd.InnerText)
            $bw.Write($data)
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- PATT ---
function Encode-PATT {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $machPatterns = @($El.SelectNodes("MachinePatterns"))

    foreach ($machPatt in $machPatterns) {
        $numPatterns = [uint16]$machPatt.GetAttribute("numPatterns")
        $numTracks = [uint16]$machPatt.GetAttribute("numTracks")
        $bw.Write($numPatterns)
        $bw.Write($numTracks)

        $patterns = @($machPatt.SelectNodes("Pattern"))
        foreach ($pat in $patterns) {
            $patName = Desanitize-XmlString $pat.GetAttribute("name")
            Write-Asciiz $bw $patName
            $bw.Write([uint16]$pat.GetAttribute("rows"))

            # Pattern data stored as base64 blob
            $base64 = $pat.InnerText
            if ($base64 -and $base64.Length -gt 0) {
                $patBytes = [Convert]::FromBase64String($base64)
                $bw.Write($patBytes)
            }
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- BLAH ---
function Encode-BLAH {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $text = $El.InnerText
    if ([string]::IsNullOrEmpty($text)) {
        $bw.Write([uint32]0)
    } else {
        $textBytes = [System.Text.Encoding]::ASCII.GetBytes($text)
        $bw.Write([uint32]$textBytes.Length)
        $bw.Write($textBytes)
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- MIDI ---
function Encode-MIDI {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $bindings = @($El.SelectNodes("Binding"))
    foreach ($bind in $bindings) {
        Write-Asciiz $bw $bind.GetAttribute("machine")
        $bw.Write([byte]$bind.GetAttribute("paramGroup"))
        $bw.Write([byte]$bind.GetAttribute("paramTrack"))
        $bw.Write([byte]$bind.GetAttribute("paramNumber"))
        $bw.Write([byte]$bind.GetAttribute("midiChannel"))
        $bw.Write([byte]$bind.GetAttribute("midiController"))
    }
    # Terminating zero byte
    $bw.Write([byte]0)

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- MACX ---
function Encode-MACX {
    param([System.Xml.XmlElement]$El)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $machines = @($El.SelectNodes("Machine"))
    $bw.Write([uint16]$machines.Count)

    foreach ($mach in $machines) {
        Write-Asciiz $bw $mach.GetAttribute("name")

        $attrs = @($mach.SelectNodes("Attribute"))
        $bw.Write([uint32]$attrs.Count)

        foreach ($attr in $attrs) {
            Write-Asciiz $bw $attr.GetAttribute("key")
            $attrSize = [uint32]$attr.GetAttribute("size")
            $bw.Write($attrSize)
            $data = [Convert]::FromBase64String($attr.InnerText)
            $bw.Write($data)
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- Raw/generic section from base64 ---
function Encode-RawSection {
    param([System.Xml.XmlElement]$El)
    $data = [Convert]::FromBase64String($El.InnerText)
    return ,$data
}

# ============================================================================
# Remap: find and replace VST/DLL paths in BMX files
# ============================================================================

function Find-VstPaths {
    # Scan binary data for embedded file paths (look for X:\ pattern)
    param([byte[]]$Bytes)

    $paths = @()
    for ($i = 0; $i -lt $Bytes.Length - 3; $i++) {
        # Look for drive letter followed by :\  (e.g., C:\)
        $b0 = $Bytes[$i]
        $b1 = $Bytes[$i + 1]
        $b2 = $Bytes[$i + 2]

        # Check: uppercase or lowercase letter, then ':\'
        $isDriveLetter = (($b0 -ge 0x41 -and $b0 -le 0x5A) -or ($b0 -ge 0x61 -and $b0 -le 0x7A))
        if ($isDriveLetter -and $b1 -eq 0x3A -and $b2 -eq 0x5C) {
            # Found a potential path - read until null terminator
            $end = $i
            while ($end -lt $Bytes.Length -and $Bytes[$end] -ne 0) { $end++ }
            $pathStr = [System.Text.Encoding]::ASCII.GetString($Bytes, $i, $end - $i)

            # Validate it looks like a real file path (contains backslashes, ends with extension)
            if ($pathStr -match '\\' -and $pathStr -match '\.\w{2,4}$') {
                $paths += [PSCustomObject]@{
                    Offset = $i
                    Length = $end - $i
                    Path   = $pathStr
                }
            }
        }
    }

    return ,$paths
}

function Invoke-RemapPaths {
    param([string]$BmxPath, [string]$OutPath, [string]$FromPrefix, [string]$ToPrefix)

    Write-Log "Starting path remap: $BmxPath"
    Write-Log "  From: $FromPrefix"
    Write-Log "  To:   $ToPrefix"

    $bytes = [System.IO.File]::ReadAllBytes($BmxPath)
    Write-Log "File size: $($bytes.Length) bytes"

    # Find all paths in the file
    $allPaths = Find-VstPaths $bytes
    Write-Log "Found $($allPaths.Count) embedded path(s) total"

    # Filter to paths that match the FromPrefix
    $matchingPaths = @($allPaths | Where-Object { $_.Path.StartsWith($FromPrefix, [StringComparison]::OrdinalIgnoreCase) })

    if ($matchingPaths.Count -eq 0) {
        Write-Log "WARNING: No paths match the prefix '$FromPrefix'"
        Write-Log "Paths found in file:"
        foreach ($p in $allPaths) {
            Write-Log "  0x$($p.Offset.ToString('X8')): $($p.Path)"
        }
        return
    }

    Write-Log "Found $($matchingPaths.Count) path(s) matching prefix"

    # Perform replacements (working backwards to preserve offsets)
    $sortedPaths = $matchingPaths | Sort-Object -Property Offset -Descending
    $replacementCount = 0

    foreach ($pathInfo in $sortedPaths) {
        $oldPath = $pathInfo.Path
        $newPath = $ToPrefix + $oldPath.Substring($FromPrefix.Length)

        # Determine available buffer space (count null bytes after the string)
        $bufferEnd = $pathInfo.Offset + $pathInfo.Length + 1  # +1 for null terminator
        while ($bufferEnd -lt $bytes.Length -and $bytes[$bufferEnd] -eq 0) {
            $bufferEnd++
        }
        $maxLen = $bufferEnd - $pathInfo.Offset - 1  # -1 for null terminator

        if ($newPath.Length -gt $maxLen) {
            Write-Log "ERROR: New path too long for buffer!"
            Write-Log "  Old: $oldPath ($($oldPath.Length) chars)"
            Write-Log "  New: $newPath ($($newPath.Length) chars)"
            Write-Log "  Buffer max: $maxLen chars"
            Write-Log "  Skipping this replacement."
            continue
        }

        Write-Log "  Replacing at 0x$($pathInfo.Offset.ToString('X8')):"
        Write-Log "    Old: $oldPath"
        Write-Log "    New: $newPath"

        # Write new path bytes
        $newPathBytes = [System.Text.Encoding]::ASCII.GetBytes($newPath)
        [Array]::Copy($newPathBytes, 0, $bytes, $pathInfo.Offset, $newPathBytes.Length)

        # Null-fill the remainder of the old string area
        $fillStart = $pathInfo.Offset + $newPathBytes.Length
        $fillEnd = $pathInfo.Offset + $pathInfo.Length
        for ($j = $fillStart; $j -le $fillEnd; $j++) {
            $bytes[$j] = 0
        }

        $replacementCount++
    }

    if ($replacementCount -gt 0) {
        [System.IO.File]::WriteAllBytes($OutPath, $bytes)
        Write-Log "Remap complete. $replacementCount path(s) replaced. Written to $OutPath"
    } else {
        Write-Log "No replacements were made."
    }
}

function Show-VstPaths {
    param([string]$BmxPath)

    $bytes = [System.IO.File]::ReadAllBytes($BmxPath)
    $allPaths = Find-VstPaths $bytes

    if ($allPaths.Count -eq 0) {
        Write-Host "No embedded file paths found in $BmxPath"
        return
    }

    Write-Host ""
    Write-Host "Embedded file paths found in:"
    Write-Host "  $BmxPath"
    Write-Host ""

    # Group by unique path
    $grouped = $allPaths | Group-Object -Property Path
    $index = 1
    foreach ($group in $grouped) {
        $count = $group.Count
        $offsets = ($group.Group | ForEach-Object { "0x$($_.Offset.ToString('X8'))" }) -join ", "
        Write-Host "  [$index] $($group.Name)"
        Write-Host "      Occurrences: $count"
        Write-Host "      Offsets: $offsets"
        Write-Host ""
        $index++
    }

    Write-Host "Total: $($allPaths.Count) path reference(s), $($grouped.Count) unique path(s)"
    Write-Host ""
}

# ============================================================================
# Main entry point
# ============================================================================

# Set up log file (use OutputFile for log path, or InputFile for list-only modes)
if ($OutputFile) {
    $LogFile = [System.IO.Path]::ChangeExtension($OutputFile, ".log")
} else {
    $LogFile = [System.IO.Path]::ChangeExtension($InputFile, ".log")
}

# Clear log file
if (Test-Path $LogFile) { Remove-Item $LogFile -Force }

try {
    switch ($Mode) {
        "decode" {
            if (-not $OutputFile) { throw "OutputFile is required for decode mode." }
            ConvertFrom-Buzz -BmxPath $InputFile -XmlPath $OutputFile
        }
        "encode" {
            if (-not $OutputFile) { throw "OutputFile is required for encode mode." }
            ConvertTo-Buzz -XmlPath $InputFile -BmxPath $OutputFile
        }
        "remap" {
            if ($ListPaths) {
                # Just list paths, no output file needed
                Show-VstPaths -BmxPath $InputFile
            } else {
                if (-not $OutputFile) { throw "OutputFile is required for remap mode (or use -ListPaths)." }
                if (-not $RemapFrom) { throw "RemapFrom is required for remap mode." }
                if (-not $RemapTo) { throw "RemapTo is required for remap mode." }
                Invoke-RemapPaths -BmxPath $InputFile -OutPath $OutputFile -FromPrefix $RemapFrom -ToPrefix $RemapTo
            }
        }
    }
} catch {
    Write-Log "ERROR: $_"
    Write-Log "Stack trace: $($_.ScriptStackTrace)"
    throw
}
