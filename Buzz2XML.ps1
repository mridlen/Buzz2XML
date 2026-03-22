<#
.SYNOPSIS
    Encode and decode Jeskola Buzz (.bmx) files to/from XML format,
    with VST path remapping support.

.DESCRIPTION
    This script can convert a Buzz .bmx binary file to an XML representation,
    convert that XML back to a .bmx binary file, remap VST plugin paths
    embedded in machine data blobs, and list or delete machines.

.PARAMETER Mode
    "decode" (bmx -> xml), "encode" (xml -> bmx), "remap" (rewrite VST paths), "machines" (list/delete machines), or "upgrade" (upgrade SampleGrid v2 to v3).

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

.PARAMETER ListMachines
    (machines mode) If set, just list all machines in the file without modifying anything.

.PARAMETER DeletePattern
    (machines mode) Wildcard pattern for machines to delete (e.g. "SVerb*").

.PARAMETER DeleteNames
    (machines mode) Array of exact machine names to delete.

.PARAMETER Help
    Show the usage/help page.

.EXAMPLE
    .\Buzz2XML.ps1 -Mode decode -InputFile test_buzz.bmx -OutputFile test_buzz.xml
    .\Buzz2XML.ps1 -Mode encode -InputFile test_buzz.xml -OutputFile test_buzz_rebuilt.bmx
    .\Buzz2XML.ps1 -Mode remap -InputFile test_buzz.bmx -OutputFile remapped.bmx -RemapFrom "C:\Program Files (x86)\Jeskola\Buzz\Gear\Vst" -RemapTo "D:\Audio\VST"
    .\Buzz2XML.ps1 -Mode remap -InputFile test_buzz.bmx -ListPaths
    .\Buzz2XML.ps1 -Mode machines -InputFile test_buzz.bmx -ListMachines
    .\Buzz2XML.ps1 -Mode machines -InputFile test_buzz.bmx -OutputFile cleaned.bmx -DeletePattern "SVerb*"
    .\Buzz2XML.ps1 -Mode machines -InputFile test_buzz.bmx -OutputFile cleaned.bmx -DeleteNames "SVerb","SVerb2"
    .\Buzz2XML.ps1 -Mode upgrade -InputFile mysong.bmx -OutputFile upgraded.bmx
    .\Buzz2XML.ps1 -Help
#>

param(
    [string]$Mode,

    [string]$InputFile,

    [string]$OutputFile,

    [string]$RemapFrom,

    [string]$RemapTo,

    [switch]$ListPaths,

    # machines mode: list or delete machines
    [switch]$ListMachines,

    [string]$DeletePattern,

    [string[]]$DeleteNames,

    # decode mode: expand PATT ParamData blobs into per-row named parameter values
    [switch]$ExpandPattData,

    [Alias("h")]
    [switch]$Help,

    # Catch-all for unrecognized flags like /?, /h, /help, --help
    [Parameter(ValueFromRemainingArguments=$true)]
    $ExtraArgs
)

# ============================================================================
# Load external modules
# ============================================================================
. "$PSScriptRoot\SampleGridUpgrade.ps1"  # SampleGridUpgrade.ps1

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
    Write-Host "  LIST / DELETE MACHINES:" -ForegroundColor Green
    Write-Host "    .\Buzz2XML.ps1 -Mode machines -InputFile <file.bmx> -ListMachines"
    Write-Host "    .\Buzz2XML.ps1 -Mode machines -InputFile <file.bmx> -OutputFile <out.bmx> ``"
    Write-Host "        -DeletePattern <wildcard>"
    Write-Host "    .\Buzz2XML.ps1 -Mode machines -InputFile <file.bmx> -OutputFile <out.bmx> ``"
    Write-Host "        -DeleteNames <name1>,<name2>,..."
    Write-Host ""
    Write-Host "    Lists or deletes machines from a Buzz song file."
    Write-Host "    Use -ListMachines to see all machines. Use -DeletePattern for wildcard"
    Write-Host "    matching (e.g. 'SVerb*') or -DeleteNames for exact names (as an array)."
    Write-Host "    Both can be combined. The Master machine cannot be deleted."
    Write-Host ""
    Write-Host "  UPGRADE SAMPLEGRID:" -ForegroundColor Green
    Write-Host "    .\Buzz2XML.ps1 -Mode upgrade -InputFile <file.bmx> -OutputFile <out.bmx>"
    Write-Host ""
    Write-Host "    Upgrades BTDSys SampleGrid 2 (v2) machines to SampleGrid 3 BETA 1 (v3)."
    Write-Host "    All 8 variants (B4/B8/B16/B32/S4/S8/S16/S32) are supported."
    Write-Host "    Parameters, patterns, and pattern editor data are migrated automatically."
    Write-Host "    Divider columns present in v2 are removed during migration."
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
    Write-Host "  -ListMachines          " -NoNewline -ForegroundColor White
    Write-Host "List all machines in the file (machines mode)"
    Write-Host "  -DeletePattern <str>   " -NoNewline -ForegroundColor White
    Write-Host "Wildcard pattern for machines to delete (machines mode)"
    Write-Host "  -DeleteNames <arr>     " -NoNewline -ForegroundColor White
    Write-Host "Exact machine names to delete, as array (machines mode)"
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
    Write-Host "  # List all machines in a song" -ForegroundColor DarkGray
    Write-Host "  .\Buzz2XML.ps1 -Mode machines -InputFile mysong.bmx -ListMachines"
    Write-Host ""
    Write-Host "  # Delete machines matching a wildcard" -ForegroundColor DarkGray
    Write-Host "  .\Buzz2XML.ps1 -Mode machines -InputFile mysong.bmx -OutputFile cleaned.bmx ``"
    Write-Host "      -DeletePattern `"SVerb*`""
    Write-Host ""
    Write-Host "  # Delete specific machines by exact name" -ForegroundColor DarkGray
    Write-Host "  .\Buzz2XML.ps1 -Mode machines -InputFile mysong.bmx -OutputFile cleaned.bmx ``"
    Write-Host "      -DeleteNames `"SVerb`",`"SVerb2`",`"SVerb22`""
    Write-Host ""
    Write-Host "  # Upgrade SampleGrid v2 machines to v3" -ForegroundColor DarkGray
    Write-Host "  .\Buzz2XML.ps1 -Mode upgrade -InputFile mysong.bmx -OutputFile upgraded.bmx"
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
if (-not $Mode -and -not $ListPaths -and -not $ListMachines -and -not $Help) { $showHelp = $true }
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
$validModes = @('decode', 'encode', 'remap', 'machines', 'upgrade')
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
    param([string]$BmxPath, [string]$XmlPath, [bool]$ExpandPattData = $false)

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
            "CONN" { Parse-CONN $bytes $sec.Offset $sec.Size $xml $root $paraInfo }  # CONN parsing (Parse-CONN)
            "CONX" { Parse-CONX $bytes $sec.Offset $sec.Size $xml $root }       # CONX parsing (Parse-CONX)
            "MACX" { Parse-MACX $bytes $sec.Offset $sec.Size $xml $root }       # MACX parsing (Parse-MACX)
            "WAVT" { Parse-WAVT $bytes $sec.Offset $sec.Size $xml $root }       # WAVT parsing (Parse-WAVT)
            "WAVE" { Parse-WAVE $bytes $sec.Offset $sec.Size $xml $root }       # WAVE parsing (Parse-WAVE)
            "PATT" { Parse-PATT $bytes $sec.Offset $sec.Size $xml $root $paraInfo $machineInputCounts $ExpandPattData }  # PATT parsing (Parse-PATT)
            "PAT2" { Parse-PAT2 $bytes $sec.Offset $sec.Size $xml $root $paraInfo }  # PAT2 parsing (Parse-PAT2)
            "PATX" { Parse-PATX $bytes $sec.Offset $sec.Size $xml $root $paraInfo }  # PATX parsing (Parse-PATX)
            "SEQU" { Parse-SEQU $bytes $sec.Offset $sec.Size $xml $root $paraInfo }  # SEQU parsing (Parse-SEQU)
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
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent, [array]$ParaInfo)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $connEl = $Xml.CreateElement("CONN")
    $Parent.AppendChild($connEl) | Out-Null

    $numConn = Read-Word $Bytes $posRef
    $connEl.SetAttribute("numConnections", $numConn)

    for ($i = 0; $i -lt $numConn; $i++) {
        $cEl = $Xml.CreateElement("Connection")
        $connEl.AppendChild($cEl) | Out-Null
        $srcIdx = Read-Word $Bytes $posRef
        $dstIdx = Read-Word $Bytes $posRef
        $cEl.SetAttribute("source", (Sanitize-XmlString $ParaInfo[$srcIdx].Name))
        $cEl.SetAttribute("destination", (Sanitize-XmlString $ParaInfo[$dstIdx].Name))
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
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent, [array]$ParaInfo, [hashtable]$MachineInputCounts, [bool]$ExpandPattData = $false)
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
        $machPattEl.SetAttribute("machine", (Sanitize-XmlString $ParaInfo[$m].Name))

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
            # Parse input connections with machine names instead of indices for safe machine deletion
            for ($ic = 0; $ic -lt $inputCount; $ic++) {
                $srcIdx = [int](Read-Word $Bytes $posRef)
                $icEl = $Xml.CreateElement("InputConnection")
                $patEl.AppendChild($icEl) | Out-Null

                # Resolve machine index to name
                if ($srcIdx -lt $ParaInfo.Count) {
                    $icEl.SetAttribute("source", (Sanitize-XmlString $ParaInfo[$srcIdx].Name))
                } else {
                    $icEl.SetAttribute("sourceIndex", $srcIdx)
                    Write-Log "  WARNING: PATT machine $m pattern $p input $ic has invalid source index $srcIdx"
                }

                # Read amp/pan data for all rows as base64
                $ampPanSize = $patLength * 4  # rows * (word amp + word pan)
                if ($ampPanSize -gt 0) {
                    $ampPanData = Read-ByteArray $Bytes $posRef $ampPanSize
                    $icEl.InnerText = [Convert]::ToBase64String($ampPanData)
                }
            }

            # Parameter data: rows * (globalSize + trackSize * numTracks)
            $paramDataSize = $paramRowSize * $patLength
            if ($paramDataSize -gt 0) {
                if (($pos + $paramDataSize) -gt $Bytes.Length) {
                    Write-Log "WARNING: PATT machine $m pattern $p would read past file end (pos=$pos, need=$paramDataSize, fileLen=$($Bytes.Length))"
                    break
                }
                $paramData = Read-ByteArray $Bytes $posRef $paramDataSize

                if ($ExpandPattData) {
                    # Expanded mode: decode into per-row elements with named parameter values
                    Expand-PattParamData $Xml $patEl $paramData $pInfo $numTracks $patLength
                } else {
                    # Default mode: store as base64 blob
                    $pdEl = $Xml.CreateElement("ParamData")
                    $patEl.AppendChild($pdEl) | Out-Null
                    $pdEl.InnerText = [Convert]::ToBase64String($paramData)
                }
            }
        }
    }
}

# --- Expand PATT ParamData into per-row named parameter elements ---
function Expand-PattParamData {
    param(
        [System.Xml.XmlDocument]$Xml,
        [System.Xml.XmlElement]$PatEl,
        [byte[]]$Data,
        [PSCustomObject]$ParaInfo,
        [int]$NumTracks,
        [int]$NumRows
    )

    $globalParams = $ParaInfo.GlobalParams
    $trackParams = $ParaInfo.TrackParams
    $globalSize = $ParaInfo.GlobalSize
    $trackSize = $ParaInfo.TrackSize
    $rowSize = $globalSize + ($trackSize * $NumTracks)

    $pdEl = $Xml.CreateElement("ParamDataExpanded")
    $PatEl.AppendChild($pdEl) | Out-Null

    for ($row = 0; $row -lt $NumRows; $row++) {
        $rowEl = $Xml.CreateElement("Row")
        $pdEl.AppendChild($rowEl) | Out-Null
        $rowEl.SetAttribute("index", $row)

        $rowOffset = $row * $rowSize

        # Global parameters
        $bytePos = 0
        foreach ($gp in $globalParams) {
            $absPos = $rowOffset + $bytePos
            if ($gp.ByteSize -eq 2) {
                $val = [BitConverter]::ToUInt16($Data, $absPos)
            } else {
                $val = [int]$Data[$absPos]
            }
            $rowEl.SetAttribute(("g_" + ($gp.Name -replace '\s+','_' -replace '[^a-zA-Z0-9_]','') + "_$bytePos"), $val)
            $bytePos += $gp.ByteSize
        }

        # Track parameters
        for ($t = 0; $t -lt $NumTracks; $t++) {
            $trackOffset = $rowOffset + $globalSize + ($t * $trackSize)
            $bytePos = 0
            foreach ($tp in $trackParams) {
                $absPos = $trackOffset + $bytePos
                if ($tp.ByteSize -eq 2) {
                    $val = [BitConverter]::ToUInt16($Data, $absPos)
                } else {
                    $val = [int]$Data[$absPos]
                }
                $rowEl.SetAttribute(("t${t}_" + ($tp.Name -replace '\s+','_' -replace '[^a-zA-Z0-9_]','') + "_$bytePos"), $val)
                $bytePos += $tp.ByteSize
            }
        }
    }
}

# --- Collapse expanded ParamData rows back into binary ---
function Collapse-PattParamData {
    param([System.Xml.XmlElement]$ExpandedEl)

    $rows = @($ExpandedEl.SelectNodes("Row"))
    if ($rows.Count -eq 0) { return [byte[]]@() }

    # Collect all attribute values in row order, sorted by attribute name
    # Attribute names encode byte position: prefix_name_bytePos
    $allBytes = [System.Collections.Generic.List[byte]]::new()
    foreach ($rowEl in $rows) {
        # Get all attributes except "index", sorted by their byte position suffix
        $attrs = @()
        foreach ($attr in $rowEl.Attributes) {
            if ($attr.Name -eq "index") { continue }
            # Extract byte position from attribute name (last segment after _)
            # Format: g_ParamName_bytePos or t0_ParamName_bytePos
            $parts = $attr.Name -split '_'
            $bytePos = [int]$parts[-1]
            $prefix = $parts[0]  # g or t0, t1, etc.

            # Determine sort key: globals first (prefix "g"), then tracks in order
            if ($prefix -eq "g") {
                $sortKey = $bytePos
            } else {
                # Track: extract track number and compute absolute position
                $trackNum = [int]($prefix.Substring(1))
                $sortKey = 100000 + ($trackNum * 1000) + $bytePos
            }

            # Determine if this is a word (2-byte) value — check if value > 255
            $val = [int]$attr.Value
            $isWord = ($val -gt 255) -or ($attr.Name -match '_Start_|_LoopStart_|_Loop_Start_|_End_|_LoopFit_|_Loop_Fit_|_Inertia_')

            $attrs += [PSCustomObject]@{
                SortKey = $sortKey
                Value = $val
                IsWord = $isWord
            }
        }

        # Sort by sort key and emit bytes
        $attrs = $attrs | Sort-Object SortKey
        foreach ($a in $attrs) {
            if ($a.IsWord) {
                $allBytes.Add([byte]($a.Value -band 0xFF))
                $allBytes.Add([byte](($a.Value -shr 8) -band 0xFF))
            } else {
                $allBytes.Add([byte]$a.Value)
            }
        }
    }

    return [byte[]]$allBytes.ToArray()
}

# --- SEQU ---
function Parse-SEQU {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent, [array]$ParaInfo)
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

        $sEl.SetAttribute("machine", (Sanitize-XmlString $ParaInfo[$machIdx].Name))
        $sEl.SetAttribute("numEvents", $numEvents)

        # bytesPerPos and bytesPerEvent are only present when numEvents > 0
        $bytesPerPos = 0
        $bytesPerEvent = 0
        if ($numEvents -gt 0) {
            $bytesPerPos = Read-Byte $Bytes $posRef
            $bytesPerEvent = Read-Byte $Bytes $posRef
            $sEl.SetAttribute("bytesPerPos", $bytesPerPos)
            $sEl.SetAttribute("bytesPerEvent", $bytesPerEvent)
        }

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

# --- PAT2 (new pattern editor data, section type TAP2 in ReBuzz) ---
function Parse-PAT2 {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent, [array]$ParaInfo)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $pat2El = $Xml.CreateElement("PAT2")
    $Parent.AppendChild($pat2El) | Out-Null

    $version = Read-Byte $Bytes $posRef
    $pat2El.SetAttribute("version", $version)

    if ($version -ne 2) {
        # Unknown version, store remainder as raw base64
        $remaining = $Size - 1
        if ($remaining -gt 0) {
            $data = Read-ByteArray $Bytes $posRef $remaining
            $pat2El.InnerText = [Convert]::ToBase64String($data)
            $pat2El.SetAttribute("encoding", "base64")
        }
        Write-Log "  PAT2: unknown version $version, stored as base64"
        return
    }

    $numMachines = $ParaInfo.Count
    for ($m = 0; $m -lt $numMachines; $m++) {
        $machEl = $Xml.CreateElement("MachineData")
        $pat2El.AppendChild($machEl) | Out-Null
        $machEl.SetAttribute("machine", (Sanitize-XmlString $ParaInfo[$m].Name))

        $numPatterns = Read-Word $Bytes $posRef
        $machEl.SetAttribute("numPatterns", $numPatterns)

        for ($p = 0; $p -lt $numPatterns; $p++) {
            $patEl = $Xml.CreateElement("Pattern")
            $machEl.AppendChild($patEl) | Out-Null

            $patName = Read-AsciizString $Bytes $posRef
            $patEl.SetAttribute("name", (Sanitize-XmlString $patName))

            $colCount = Read-Int32 $Bytes $posRef
            $patEl.SetAttribute("columnCount", $colCount)

            for ($c = 0; $c -lt $colCount; $c++) {
                $colEl = $Xml.CreateElement("Column")
                $patEl.AppendChild($colEl) | Out-Null

                $pMachIdx = [int](Read-Word $Bytes $posRef)
                # Store machine name instead of index (0xFFFF = no machine)
                if ($pMachIdx -eq 0xFFFF) {
                    $colEl.SetAttribute("targetMachine", "none")
                } elseif ($pMachIdx -lt $ParaInfo.Count) {
                    $colEl.SetAttribute("targetMachine", (Sanitize-XmlString $ParaInfo[$pMachIdx].Name))
                } else {
                    $colEl.SetAttribute("targetMachineIndex", $pMachIdx)
                    Write-Log "  WARNING: PAT2 machine $m pattern $p column $c has invalid target index $pMachIdx"
                }

                $group = Read-Int32 $Bytes $posRef
                $indexInGroup = Read-Int32 $Bytes $posRef
                $track = Read-Int32 $Bytes $posRef
                $colEl.SetAttribute("group", $group)
                $colEl.SetAttribute("indexInGroup", $indexInGroup)
                $colEl.SetAttribute("track", $track)

                $numEvents = Read-Int32 $Bytes $posRef
                $colEl.SetAttribute("numEvents", $numEvents)

                for ($e = 0; $e -lt $numEvents; $e++) {
                    $evEl = $Xml.CreateElement("Event")
                    $colEl.AppendChild($evEl) | Out-Null
                    $evEl.SetAttribute("time", (Read-Int32 $Bytes $posRef))
                    $evEl.SetAttribute("value", (Read-Int32 $Bytes $posRef))
                    $evEl.SetAttribute("duration", (Read-Int32 $Bytes $posRef))
                }

                $numMeta = Read-Int32 $Bytes $posRef
                for ($md = 0; $md -lt $numMeta; $md++) {
                    $metaEl = $Xml.CreateElement("Meta")
                    $colEl.AppendChild($metaEl) | Out-Null
                    $metaEl.SetAttribute("key", (Read-AsciizString $Bytes $posRef))
                    $metaEl.SetAttribute("value", (Read-AsciizString $Bytes $posRef))
                }
            }
        }
    }
    Write-Log "  PAT2: parsed $numMachines machines (version $version)"
}

# --- PATX (pattern editor assignment, section type XTAP in ReBuzz) ---
function Parse-PATX {
    param([byte[]]$Bytes, [uint32]$Offset, [uint32]$Size, [System.Xml.XmlDocument]$Xml, [System.Xml.XmlElement]$Parent, [array]$ParaInfo)
    $pos = [int]$Offset
    $posRef = [ref]$pos

    $patxEl = $Xml.CreateElement("PATX")
    $Parent.AppendChild($patxEl) | Out-Null

    $version = Read-Byte $Bytes $posRef
    $patxEl.SetAttribute("version", $version)

    if ($version -ne 1) {
        # Unknown version, store remainder as raw base64
        $remaining = $Size - 1
        if ($remaining -gt 0) {
            $data = Read-ByteArray $Bytes $posRef $remaining
            $patxEl.InnerText = [Convert]::ToBase64String($data)
            $patxEl.SetAttribute("encoding", "base64")
        }
        Write-Log "  PATX: unknown version $version, stored as base64"
        return
    }

    $numMachines = $ParaInfo.Count
    for ($m = 0; $m -lt $numMachines; $m++) {
        $machEl = $Xml.CreateElement("MachineEditor")
        $patxEl.AppendChild($machEl) | Out-Null
        $machEl.SetAttribute("machine", (Sanitize-XmlString $ParaInfo[$m].Name))

        $numPatterns = Read-Word $Bytes $posRef
        $machEl.SetAttribute("numPatterns", $numPatterns)

        for ($p = 0; $p -lt $numPatterns; $p++) {
            $peEl = $Xml.CreateElement("PatternEditor")
            $machEl.AppendChild($peEl) | Out-Null

            $editorIdx = [int](Read-Word $Bytes $posRef)
            # Store machine name instead of index (0xFFFF = built-in editor)
            if ($editorIdx -eq 0xFFFF) {
                $peEl.SetAttribute("editor", "builtin")
            } elseif ($editorIdx -lt $ParaInfo.Count) {
                $peEl.SetAttribute("editor", (Sanitize-XmlString $ParaInfo[$editorIdx].Name))
            } else {
                $peEl.SetAttribute("editorIndex", $editorIdx)
                Write-Log "  WARNING: PATX machine $m pattern $p has invalid editor index $editorIdx"
            }
        }
    }
    Write-Log "  PATX: parsed $numMachines machines (version $version)"
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

    # Build machine name -> index lookup from PARA section
    $machineNameToIndex = @{}
    $paraEl = $root.SelectSingleNode("PARA")
    if ($paraEl) {
        $paraMachines = @($paraEl.SelectNodes("Machine"))
        for ($mi = 0; $mi -lt $paraMachines.Count; $mi++) {
            $mName = $paraMachines[$mi].GetAttribute("name")
            $machineNameToIndex[$mName] = $mi
        }
    }

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
            "CONN" { $data = Encode-CONN $child $machineNameToIndex }  # CONN encoding (Encode-CONN)
            "SEQU" { $data = Encode-SEQU $child $machineNameToIndex }  # SEQU encoding (Encode-SEQU)
            "WAVT" { $data = Encode-WAVT $child }           # WAVT encoding (Encode-WAVT)
            "WAVE" { $data = Encode-WAVE $child }           # WAVE encoding (Encode-WAVE)
            "PATT" { $data = Encode-PATT $child $machineNameToIndex }  # PATT encoding (Encode-PATT)
            "BLAH" { $data = Encode-BLAH $child }           # BLAH encoding (Encode-BLAH)
            "MIDI" { $data = Encode-MIDI $child }           # MIDI encoding (Encode-MIDI)
            "MACX" { $data = Encode-MACX $child }           # MACX encoding (Encode-MACX)
            "PAT2" { $data = Encode-PAT2 $child $machineNameToIndex }  # PAT2 encoding (Encode-PAT2)
            "PATX" { $data = Encode-PATX $child $machineNameToIndex }  # PATX encoding (Encode-PATX)
            default {
                # Base64-encoded raw section (CONX, PDLG, BGUI, etc.)
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
    param([System.Xml.XmlElement]$El, [hashtable]$NameToIndex)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $conns = @($El.SelectNodes("Connection"))
    $bw.Write([uint16]$conns.Count)

    foreach ($conn in $conns) {
        $srcName = $conn.GetAttribute("source")
        $dstName = $conn.GetAttribute("destination")
        $bw.Write([uint16]$NameToIndex[$srcName])
        $bw.Write([uint16]$NameToIndex[$dstName])
        $bw.Write([uint16]$conn.GetAttribute("amp"))
        $bw.Write([uint16]$conn.GetAttribute("pan"))
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- SEQU ---
function Encode-SEQU {
    param([System.Xml.XmlElement]$El, [hashtable]$NameToIndex)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $bw.Write([uint32]$El.GetAttribute("endOfSong"))
    $bw.Write([uint32]$El.GetAttribute("beginLoop"))
    $bw.Write([uint32]$El.GetAttribute("endLoop"))

    $seqs = @($El.SelectNodes("Sequence"))
    $bw.Write([uint16]$seqs.Count)

    foreach ($seq in $seqs) {
        $machName = $seq.GetAttribute("machine")
        $bw.Write([uint16]$NameToIndex[$machName])
        $numEvents = [uint32]$seq.GetAttribute("numEvents")
        $bw.Write($numEvents)

        # bytesPerPos and bytesPerEvent are only written when numEvents > 0
        $bytesPerPos = 0
        $bytesPerEvent = 0
        if ($numEvents -gt 0) {
            $bytesPerPos = [byte]$seq.GetAttribute("bytesPerPos")
            $bytesPerEvent = [byte]$seq.GetAttribute("bytesPerEvent")
            $bw.Write($bytesPerPos)
            $bw.Write($bytesPerEvent)
        }

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
    param([System.Xml.XmlElement]$El, [hashtable]$NameToIndex)
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

            # Input connection data: machine name -> index + amp/pan blob
            $inputConns = @($pat.SelectNodes("InputConnection"))
            foreach ($ic in $inputConns) {
                # Resolve machine name back to index
                $srcName = $ic.GetAttribute("source")
                if ($srcName) {
                    $bw.Write([uint16]$NameToIndex[$srcName])
                } else {
                    # Fallback: use stored numeric index
                    $bw.Write([uint16]$ic.GetAttribute("sourceIndex"))
                }
                # Write amp/pan data
                $ampPanBase64 = $ic.InnerText
                if ($ampPanBase64 -and $ampPanBase64.Length -gt 0) {
                    $ampPanBytes = [Convert]::FromBase64String($ampPanBase64)
                    $bw.Write($ampPanBytes)
                }
            }

            # Parameter data blob — supports both base64 (ParamData) and expanded (ParamDataExpanded)
            $paramDataEl = $pat.SelectSingleNode("ParamData")
            $paramDataExpandedEl = $pat.SelectSingleNode("ParamDataExpanded")
            if ($paramDataEl) {
                $paramBytes = [Convert]::FromBase64String($paramDataEl.InnerText)
                $bw.Write($paramBytes)
            } elseif ($paramDataExpandedEl) {
                # Re-pack expanded rows back into binary
                $paramBytes = Collapse-PattParamData $paramDataExpandedEl
                $bw.Write($paramBytes)
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

# --- PAT2 ---
function Encode-PAT2 {
    param([System.Xml.XmlElement]$El, [hashtable]$NameToIndex)
    # Check if this was stored as raw base64 (unknown version)
    if ($El.GetAttribute("encoding") -eq "base64") {
        return ,([Convert]::FromBase64String($El.InnerText))
    }

    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $version = [byte]$El.GetAttribute("version")
    $bw.Write($version)

    $machines = @($El.SelectNodes("MachineData"))
    foreach ($machEl in $machines) {
        $patterns = @($machEl.SelectNodes("Pattern"))
        $bw.Write([uint16]$patterns.Count)

        foreach ($pat in $patterns) {
            Write-Asciiz $bw ($pat.GetAttribute("name"))

            $columns = @($pat.SelectNodes("Column"))
            $bw.Write([int32]$columns.Count)

            foreach ($col in $columns) {
                # Resolve target machine name to index
                $targetMach = $col.GetAttribute("targetMachine")
                if ($targetMach -eq "none") {
                    $bw.Write([uint16]0xFFFF)
                } elseif ($targetMach -and $NameToIndex.ContainsKey($targetMach)) {
                    $bw.Write([uint16]$NameToIndex[$targetMach])
                } else {
                    # Fallback: use stored numeric index
                    $bw.Write([uint16]$col.GetAttribute("targetMachineIndex"))
                }

                $bw.Write([int32]$col.GetAttribute("group"))
                $bw.Write([int32]$col.GetAttribute("indexInGroup"))
                $bw.Write([int32]$col.GetAttribute("track"))

                $events = @($col.SelectNodes("Event"))
                $bw.Write([int32]$events.Count)
                foreach ($ev in $events) {
                    $bw.Write([int32]$ev.GetAttribute("time"))
                    $bw.Write([int32]$ev.GetAttribute("value"))
                    $bw.Write([int32]$ev.GetAttribute("duration"))
                }

                $metas = @($col.SelectNodes("Meta"))
                $bw.Write([int32]$metas.Count)
                foreach ($meta in $metas) {
                    Write-Asciiz $bw ($meta.GetAttribute("key"))
                    Write-Asciiz $bw ($meta.GetAttribute("value"))
                }
            }
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
}

# --- PATX ---
function Encode-PATX {
    param([System.Xml.XmlElement]$El, [hashtable]$NameToIndex)
    # Check if this was stored as raw base64 (unknown version)
    if ($El.GetAttribute("encoding") -eq "base64") {
        return ,([Convert]::FromBase64String($El.InnerText))
    }

    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)

    $version = [byte]$El.GetAttribute("version")
    $bw.Write($version)

    $machines = @($El.SelectNodes("MachineEditor"))
    foreach ($machEl in $machines) {
        $editors = @($machEl.SelectNodes("PatternEditor"))
        $bw.Write([uint16]$editors.Count)

        foreach ($pe in $editors) {
            $editor = $pe.GetAttribute("editor")
            if ($editor -eq "builtin") {
                $bw.Write([uint16]0xFFFF)
            } elseif ($editor -and $NameToIndex.ContainsKey($editor)) {
                $bw.Write([uint16]$NameToIndex[$editor])
            } else {
                # Fallback: use stored numeric index
                $bw.Write([uint16]$pe.GetAttribute("editorIndex"))
            }
        }
    }

    $result = $ms.ToArray()
    $bw.Close(); $ms.Close()
    return ,$result
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
# Machines: list and delete machines from BMX files
# ============================================================================

# Show-MachineList: list all machines in a BMX file
function Show-MachineList {
    param([string]$BmxPath)

    Write-Log "Listing machines in: $BmxPath"

    # Decode to XML in memory
    $bytes = [System.IO.File]::ReadAllBytes($BmxPath)
    Write-Log "File size: $($bytes.Length) bytes"

    # Read header and find PARA section
    $magic = [System.Text.Encoding]::ASCII.GetString($bytes, 0, 4)
    if ($magic -ne "Buzz") { throw "Not a valid Buzz file (magic='$magic')" }
    $numSections = [BitConverter]::ToUInt32($bytes, 4)

    $paraOffset = $null
    $paraSize = $null
    for ($i = 0; $i -lt $numSections; $i++) {
        $dirPos = 8 + ($i * 12)
        $secName = [System.Text.Encoding]::ASCII.GetString($bytes, $dirPos, 4).TrimEnd([char]0)
        $secOffset = [BitConverter]::ToUInt32($bytes, $dirPos + 4)
        $secSize = [BitConverter]::ToUInt32($bytes, $dirPos + 8)
        if ($secName -eq "PARA") {
            $paraOffset = $secOffset
            $paraSize = $secSize
            break
        }
    }

    if ($null -eq $paraOffset) {
        Write-Host "No PARA section found in $BmxPath"
        return
    }

    # Parse machine names from PARA
    $pos = [int]$paraOffset
    $posRef = [ref]$pos
    $numMachines = Read-DWord $bytes $posRef

    Write-Host ""
    Write-Host "Machines in: $BmxPath"
    Write-Host ""

    for ($m = 0; $m -lt $numMachines; $m++) {
        $name = Read-AsciizString $bytes $posRef
        $type = Read-AsciizString $bytes $posRef
        $numGlobal = Read-DWord $bytes $posRef
        $numTrack = Read-DWord $bytes $posRef
        $totalParams = $numGlobal + $numTrack

        # Skip parameter definitions
        for ($p = 0; $p -lt $totalParams; $p++) {
            $null = Read-Byte $bytes $posRef      # type
            $null = Read-AsciizString $bytes $posRef  # name
            $null = Read-DWord $bytes $posRef      # minValue
            $null = Read-DWord $bytes $posRef      # maxValue
            $null = Read-DWord $bytes $posRef      # noValue
            $null = Read-DWord $bytes $posRef      # flags
            $null = Read-DWord $bytes $posRef      # defValue
        }

        # Determine machine kind
        $isHidden = ($name.Length -gt 0) -and ([byte][char]$name[0] -eq 1)
        $kind = if ($isHidden) { "(hidden)" } else { $type }

        Write-Host "  [$($m + 1)] $name  [$kind]"
    }

    Write-Host ""
    Write-Host "Total: $numMachines machine(s)"
    Write-Host ""
}

# Invoke-DeleteMachines: remove machines from a BMX file
function Invoke-DeleteMachines {
    param(
        [string]$BmxPath,
        [string]$OutPath,
        [string]$Pattern,       # wildcard pattern (e.g. "SVerb*")
        [string[]]$Names        # exact names to delete
    )

    Write-Log "Starting machine delete: $BmxPath -> $OutPath"

    # Build the set of names to delete by first decoding to XML in memory
    # Step 1: Decode to temp XML file
    $tempXml = [System.IO.Path]::GetTempFileName()
    $tempXml = [System.IO.Path]::ChangeExtension($tempXml, ".xml")
    Write-Log "  Decoding to temp XML: $tempXml"
    ConvertFrom-Buzz -BmxPath $BmxPath -XmlPath $tempXml  # decode (ConvertFrom-Buzz)

    # Step 2: Load XML and find machines to delete
    $xml = New-Object System.Xml.XmlDocument
    $xml.Load($tempXml)
    $root = $xml.DocumentElement

    $paraEl = $root.SelectSingleNode("PARA")
    if (-not $paraEl) { throw "No PARA section found in decoded XML" }

    $allMachineNames = @()
    foreach ($mach in @($paraEl.SelectNodes("Machine"))) {
        $allMachineNames += $mach.GetAttribute("name")
    }

    # Build delete set from pattern and/or exact names
    $deleteSet = @{}
    if ($Pattern) {
        foreach ($name in $allMachineNames) {
            if ($name -like $Pattern) {
                $deleteSet[$name] = $true
            }
        }
        Write-Log "  Pattern '$Pattern' matched: $($deleteSet.Count) machine(s)"
    }
    if ($Names) {
        foreach ($n in $Names) {
            if ($n -in $allMachineNames) {
                $deleteSet[$n] = $true
            } else {
                Write-Log "  WARNING: Machine '$n' not found in file, skipping"
                Write-Host "WARNING: Machine '$n' not found in file, skipping"
            }
        }
    }

    if ($deleteSet.Count -eq 0) {
        Write-Host "No machines matched for deletion."
        Write-Log "No machines matched for deletion."
        Remove-Item $tempXml -Force -ErrorAction SilentlyContinue
        return
    }

    # Prevent deleting the Master machine (index 0)
    if ($deleteSet.ContainsKey("Master")) {
        Write-Host "WARNING: Cannot delete the Master machine, skipping"
        Write-Log "  WARNING: Cannot delete the Master machine, skipping"
        $deleteSet.Remove("Master")
    }

    # Auto-detect associated hidden pattern editor (pe) machines
    # Pattern: if a machine is followed in PARA order by a hidden machine (name starts with 0x01),
    # that hidden machine is its Pattern XP editor and should be deleted along with it.
    # Note: in XML, the 0x01 byte is sanitized to "&#x01;" since it's not valid XML 1.0,
    # and the XML parser preserves it as literal text, so we check for both forms.
    $peCount = 0
    for ($i = 0; $i -lt $allMachineNames.Count; $i++) {
        $name = $allMachineNames[$i]
        if ($deleteSet.ContainsKey($name)) {
            # Check if the next machine(s) are hidden pe editors for this machine
            $j = $i + 1
            while ($j -lt $allMachineNames.Count) {
                $nextName = $allMachineNames[$j]
                $isHidden = $nextName.StartsWith("&#x01;") -or (($nextName.Length -gt 0) -and ([byte][char]$nextName[0] -eq 1))
                if ($isHidden -and -not $deleteSet.ContainsKey($nextName)) {
                    $deleteSet[$nextName] = $true
                    $peCount++
                    Write-Log "  Auto-including pattern editor '$nextName' (follows '$name')"
                }
                if (-not $isHidden) { break }  # stop at next non-hidden machine
                $j++
            }
        }
    }
    if ($peCount -gt 0) {
        Write-Host "Also removing $peCount associated pattern editor(s)"
    }

    Write-Host ""
    Write-Host "Deleting $($deleteSet.Count) machine(s):"
    foreach ($name in $deleteSet.Keys) {
        $isHidden = ($name.Length -gt 0) -and ([byte][char]$name[0] -eq 1)
        $label = if ($isHidden) { "$name (pattern editor)" } else { $name }
        Write-Host "  - $label"
    }
    Write-Host ""

    # Step 3: Remove machines from each section

    # --- PARA: remove machine elements ---
    $nodesToRemove = @()
    foreach ($mach in @($paraEl.SelectNodes("Machine"))) {
        if ($deleteSet.ContainsKey($mach.GetAttribute("name"))) {
            $nodesToRemove += $mach
        }
    }
    foreach ($node in $nodesToRemove) { $paraEl.RemoveChild($node) | Out-Null }
    # Update machine count attribute if present
    $paraEl.SetAttribute("numMachines", @($paraEl.SelectNodes("Machine")).Count)
    Write-Log "  PARA: removed $($nodesToRemove.Count) machine(s)"

    # --- MACH: remove machine elements ---
    $machEl = $root.SelectSingleNode("MACH")
    if ($machEl) {
        $nodesToRemove = @()
        foreach ($mach in @($machEl.SelectNodes("Machine"))) {
            if ($deleteSet.ContainsKey($mach.GetAttribute("name"))) {
                $nodesToRemove += $mach
            }
        }
        foreach ($node in $nodesToRemove) { $machEl.RemoveChild($node) | Out-Null }
        $machEl.SetAttribute("numMachines", @($machEl.SelectNodes("Machine")).Count)
        Write-Log "  MACH: removed $($nodesToRemove.Count) machine(s)"
    }

    # --- CONN: remove connections that reference deleted machines ---
    $connEl = $root.SelectSingleNode("CONN")
    if ($connEl) {
        $nodesToRemove = @()
        foreach ($conn in @($connEl.SelectNodes("Connection"))) {
            $src = $conn.GetAttribute("source")
            $dst = $conn.GetAttribute("destination")
            if ($deleteSet.ContainsKey($src) -or $deleteSet.ContainsKey($dst)) {
                $nodesToRemove += $conn
            }
        }
        foreach ($node in $nodesToRemove) { $connEl.RemoveChild($node) | Out-Null }
        $connEl.SetAttribute("numConnections", @($connEl.SelectNodes("Connection")).Count)
        Write-Log "  CONN: removed $($nodesToRemove.Count) connection(s)"
    }

    # --- CONX: drop entirely (contains machine indices that shift after deletion) ---
    $conxEl = $root.SelectSingleNode("CONX")
    if ($conxEl) {
        $root.RemoveChild($conxEl) | Out-Null
        Write-Log "  CONX: removed (machine indices would be invalidated)"
    }

    # --- PATT: remove MachinePatterns for deleted machines ---
    $pattEl = $root.SelectSingleNode("PATT")
    if ($pattEl) {
        $nodesToRemove = @()
        foreach ($mp in @($pattEl.SelectNodes("MachinePatterns"))) {
            if ($deleteSet.ContainsKey($mp.GetAttribute("machine"))) {
                $nodesToRemove += $mp
            }
        }
        foreach ($node in $nodesToRemove) { $pattEl.RemoveChild($node) | Out-Null }
        Write-Log "  PATT: removed $($nodesToRemove.Count) machine pattern set(s)"
    }

    # --- SEQU: remove sequences for deleted machines ---
    $sequEl = $root.SelectSingleNode("SEQU")
    if ($sequEl) {
        $nodesToRemove = @()
        foreach ($seq in @($sequEl.SelectNodes("Sequence"))) {
            if ($deleteSet.ContainsKey($seq.GetAttribute("machine"))) {
                $nodesToRemove += $seq
            }
        }
        foreach ($node in $nodesToRemove) { $sequEl.RemoveChild($node) | Out-Null }
        $remaining = @($sequEl.SelectNodes("Sequence")).Count
        $sequEl.SetAttribute("numSequences", $remaining)
        Write-Log "  SEQU: removed $($nodesToRemove.Count) sequence(s), $remaining remaining"
    }

    # --- MACX: remove machine elements for deleted machines ---
    $macxEl = $root.SelectSingleNode("MACX")
    if ($macxEl) {
        $nodesToRemove = @()
        foreach ($mach in @($macxEl.SelectNodes("Machine"))) {
            if ($deleteSet.ContainsKey($mach.GetAttribute("name"))) {
                $nodesToRemove += $mach
            }
        }
        foreach ($node in $nodesToRemove) { $macxEl.RemoveChild($node) | Out-Null }
        $macxEl.SetAttribute("numMachines", @($macxEl.SelectNodes("Machine")).Count)
        Write-Log "  MACX: removed $($nodesToRemove.Count) machine(s)"
    }

    # --- MIDI: remove bindings for deleted machines ---
    $midiEl = $root.SelectSingleNode("MIDI")
    if ($midiEl) {
        $nodesToRemove = @()
        foreach ($bind in @($midiEl.SelectNodes("Binding"))) {
            if ($deleteSet.ContainsKey($bind.GetAttribute("machine"))) {
                $nodesToRemove += $bind
            }
        }
        foreach ($node in $nodesToRemove) { $midiEl.RemoveChild($node) | Out-Null }
        Write-Log "  MIDI: removed $($nodesToRemove.Count) binding(s)"
    }

    # --- PDLG: drop (contains machine indices for dialog positions) ---
    $pdlgEl = $root.SelectSingleNode("PDLG")
    if ($pdlgEl) {
        $root.RemoveChild($pdlgEl) | Out-Null
        Write-Log "  PDLG: removed (machine indices would be invalidated)"
    }

    # --- PAT2: remove MachineData for deleted machines ---
    $pat2El = $root.SelectSingleNode("PAT2")
    if ($pat2El -and -not $pat2El.GetAttribute("encoding")) {
        $nodesToRemove = @()
        foreach ($md in @($pat2El.SelectNodes("MachineData"))) {
            if ($deleteSet.ContainsKey($md.GetAttribute("machine"))) {
                $nodesToRemove += $md
            }
        }
        foreach ($node in $nodesToRemove) { $pat2El.RemoveChild($node) | Out-Null }
        Write-Log "  PAT2: removed $($nodesToRemove.Count) machine data set(s)"
    } elseif ($pat2El) {
        # Raw base64 PAT2 (unknown version) — must drop
        $root.RemoveChild($pat2El) | Out-Null
        Write-Log "  PAT2: removed (raw base64, cannot fix machine indices)"
    }

    # --- PATX: remove MachineEditor for deleted machines ---
    $patxEl = $root.SelectSingleNode("PATX")
    if ($patxEl -and -not $patxEl.GetAttribute("encoding")) {
        $nodesToRemove = @()
        foreach ($me in @($patxEl.SelectNodes("MachineEditor"))) {
            if ($deleteSet.ContainsKey($me.GetAttribute("machine"))) {
                $nodesToRemove += $me
            }
        }
        foreach ($node in $nodesToRemove) { $patxEl.RemoveChild($node) | Out-Null }
        Write-Log "  PATX: removed $($nodesToRemove.Count) machine editor assignment(s)"
    } elseif ($patxEl) {
        # Raw base64 PATX (unknown version) — must drop
        $root.RemoveChild($patxEl) | Out-Null
        Write-Log "  PATX: removed (raw base64, cannot fix machine indices)"
    }

    # (Legacy placeholder for future indexed sections)
    foreach ($secName in @()) {
        $secEl = $root.SelectSingleNode($secName)
        if ($secEl) {
            $root.RemoveChild($secEl) | Out-Null
            Write-Log "  ${secName}: removed (machine indices would be invalidated)"
        }
    }

    # Step 4: Save modified XML and re-encode to BMX
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.IndentChars = "  "
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.CheckCharacters = $false

    $writer = [System.Xml.XmlWriter]::Create($tempXml, $settings)
    $xml.Save($writer)
    $writer.Close()
    Write-Log "  Saved modified XML to $tempXml"

    ConvertTo-Buzz -XmlPath $tempXml -BmxPath $OutPath  # encode (ConvertTo-Buzz)

    # Clean up temp file
    Remove-Item $tempXml -Force -ErrorAction SilentlyContinue

    Write-Host "Done! $($deleteSet.Count) machine(s) deleted."
    Write-Host "Output written to: $OutPath"
    Write-Log "Machine delete complete. Output: $OutPath"
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
            ConvertFrom-Buzz -BmxPath $InputFile -XmlPath $OutputFile -ExpandPattData $ExpandPattData.IsPresent
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
        "machines" {
            if ($ListMachines) {
                # Just list machines, no output file needed
                Show-MachineList -BmxPath $InputFile  # machine listing (Show-MachineList)
            } else {
                if (-not $OutputFile) { throw "OutputFile is required for machines delete mode (or use -ListMachines)." }
                if (-not $DeletePattern -and -not $DeleteNames) { throw "DeletePattern and/or DeleteNames is required for machines delete mode." }
                Invoke-DeleteMachines -BmxPath $InputFile -OutPath $OutputFile -Pattern $DeletePattern -Names $DeleteNames  # machine delete (Invoke-DeleteMachines)
            }
        }
        "upgrade" {
            if (-not $OutputFile) { throw "OutputFile is required for upgrade mode." }
            Invoke-UpgradeSampleGrid -BmxPath $InputFile -OutPath $OutputFile  # SampleGrid upgrade (Invoke-UpgradeSampleGrid in SampleGridUpgrade.ps1)
        }
    }
} catch {
    Write-Log "ERROR: $_"
    Write-Log "Stack trace: $($_.ScriptStackTrace)"
    throw
}
