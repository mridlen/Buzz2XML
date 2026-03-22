# ============================================================================
# SampleGrid Upgrade - Upgrade v1/v2 (BTDSys) to v3 (SampleGrid 3 BETA 1)
# ============================================================================

# ============================================================================
# Version detection and variant classification
# ============================================================================
# SampleGrid comes in multiple versions/revisions with different param layouts.
# Known versions (identified by DLL name / PARA type):
#   v1: "BTDSys SampleGrid - switch x N" or "BTDSys SampleGrid - byte x N"
#       (9 track params, wavetable samples only, no drumkits)
#   v2 older 050422: "BTDSys SampleGrid 2 BETA 050422" (15 track params)
#   v2 mid 050604: "BTDSys SampleGrid 2 BETA 050604" (20 track params)
#   v2 late 050716: "BTDSys SampleGrid 2 BETA 050716" (26 track params)
#   v2 061110: "BTDSys SampleGrid 2 BETA 061110" (31 track params)
# All get upgraded to "SampleGrid 3 BETA 1".
# ============================================================================

# Detect if a type string is ANY upgradeable SampleGrid variant (v1 or v2)
function Test-SGridUpgradeable {
    param([string]$TypeString)
    return (($TypeString -like "BTDSys SampleGrid 2*") -or ($TypeString -like "BTDSys SampleGrid - *"))
}

# Legacy alias for backward compatibility
function Test-SGridV2 {
    param([string]$TypeString)
    return ($TypeString -like "BTDSys SampleGrid 2*")
}

# Detect if a type string is a v1 SampleGrid
function Test-SGridV1 {
    param([string]$TypeString)
    return ($TypeString -like "BTDSys SampleGrid - *")
}

# Detect if type is a B (byte trigger) or S (switch trigger) variant
# v1: "BTDSys SampleGrid - byte x N" or "BTDSys SampleGrid - switch x N"
# v2: "BTDSys SampleGrid 2 - BN" or "BTDSys SampleGrid 2 - SN"
function Test-SGridBVariant {
    param([string]$TypeString)
    if ($TypeString -like "BTDSys SampleGrid - byte*") { return $true }
    return ($TypeString -match '- B\d+$')
}

# Extract channel count from type string
# v2: "BTDSys SampleGrid 2 - S4" -> 4
# v1: "BTDSys SampleGrid - switch x 8" -> 8
function Get-SGridChannelCount {
    param([string]$TypeString)
    # v2 format: ends with S4, B16, etc.
    if ($TypeString -match '[BS](\d+)$') {
        return [int]$Matches[1]
    }
    # v1 format: "- switch x 8" or "- byte x 16"
    if ($TypeString -match 'x\s+(\d+)$') {
        return [int]$Matches[1]
    }
    return 0
}

# Classify the revision based on type string and numTrackParams from PARA.
# Returns: "v1switch", "v1byte", "older", "mid", "late", "061110"
function Get-SGridRevision {
    param(
        [string]$TypeString,
        [int]$NumTrackParams = 0
    )
    # v1 detection by type string
    if ($TypeString -like "BTDSys SampleGrid - switch*") { return "v1switch" }
    if ($TypeString -like "BTDSys SampleGrid - byte*") { return "v1byte" }

    # v2 detection: 061110 by type string
    if ($TypeString -like "*061110*") { return "061110" }

    # Other v2 revisions: classify by numTrackParams
    switch ($NumTrackParams) {
        15 { return "older" }
        20 { return "mid" }
        26 { return "late" }
        31 { return "061110" }  # fallback if type string didn't have 061110
        default {
            Write-Log "    WARNING: Unknown v2 revision with $NumTrackParams track params, treating as older"
            return "older"
        }
    }
}

# Build v3 DLL name from any v1/v2 DLL name
# v2: "BTDSys SampleGrid 2 BETA 050422 - S04" -> "SampleGrid 3 BETA 1 - S04"
# v1: "BTDSys SampleGrid - switch x 8" -> "SampleGrid 3 BETA 1 - S08"
# v1: "BTDSys SampleGrid - byte x 16" -> "SampleGrid 3 BETA 1 - B16"
function Get-SGridV3DllName {
    param([string]$V2Dll)
    # v2 format: ends with - S04, - B16, etc.
    if ($V2Dll -match '- ([BS]\d+)$') {
        return "SampleGrid 3 BETA 1 - $($Matches[1])"
    }
    # v1 format: "- switch x N" or "- byte x N"
    if ($V2Dll -match '(switch|byte)\s+x\s+(\d+)$') {
        $variant = if ($Matches[1] -eq "byte") { "B" } else { "S" }
        $count = $Matches[2].PadLeft(2, '0')
        return "SampleGrid 3 BETA 1 - $variant$count"
    }
    return $null
}

# Build v3 PARA type name from any v1/v2 PARA type
# v2: "BTDSys SampleGrid 2 - S4" -> "SampleGrid 3 BETA 1 - S4"
# v1: "BTDSys SampleGrid - switch x 8" -> "SampleGrid 3 BETA 1 - S8"
# v1: "BTDSys SampleGrid - byte x 16" -> "SampleGrid 3 BETA 1 - B16"
function Get-SGridV3TypeName {
    param([string]$V2Type)
    # v2 format
    if ($V2Type -match '- ([BS]\d+)$') {
        return "SampleGrid 3 BETA 1 - $($Matches[1])"
    }
    # v1 format
    if ($V2Type -match '(switch|byte)\s+x\s+(\d+)$') {
        $variant = if ($Matches[1] -eq "byte") { "B" } else { "S" }
        return "SampleGrid 3 BETA 1 - $variant$($Matches[2])"
    }
    return $null
}

# v2 (061110) track param divider positions (0-based indices among the 31 track params)
$script:SGridV2TrackDividerIndices = @(1, 10, 17, 23, 28)

# v2 global param divider positions (0-based indices)
# Position depends on channel count: First Wave at 0, Divider at 1, Triggers..., Divider at (1+numChannels+1)
function Get-SGridV2GlobalDividerIndices {
    param([int]$NumChannels)
    return @(1, (2 + $NumChannels))
}

# ============================================================================
# Build v3 PARA parameters from v2 type info
# ============================================================================
function Build-SGridV3ParaParams {
    param(
        [System.Xml.XmlDocument]$Xml,
        [string]$V2Type  # PARA type string like "BTDSys SampleGrid 2 BETA 061110 - B4"
    )

    $numChannels = Get-SGridChannelCount $V2Type
    $isBVariant = Test-SGridBVariant $V2Type

    # v3 trigger type: B variants use type=2 (byte), S variants use type=1 (switch)
    $trigType = if ($isBVariant) { "2" } else { "1" }

    $params = @()

    # v3 global order: Trigger 0..N-1, First Wave, TrigType...Inertia
    # Triggers
    for ($i = 0; $i -lt $numChannels; $i++) {
        $p = $Xml.CreateElement("Parameter")
        $p.SetAttribute("scope", "global")
        $p.SetAttribute("type", $trigType)
        $p.SetAttribute("name", "Trigger $i")
        $p.SetAttribute("minValue", "0")
        $p.SetAttribute("maxValue", "254")
        $p.SetAttribute("noValue", "255")
        $p.SetAttribute("flags", "0")
        $p.SetAttribute("defValue", "0")
        $params += $p
    }

    # First Wave
    $p = $Xml.CreateElement("Parameter"); $p.SetAttribute("scope", "global"); $p.SetAttribute("type", "2")
    $p.SetAttribute("name", "First Wave"); $p.SetAttribute("minValue", "1"); $p.SetAttribute("maxValue", "200")
    $p.SetAttribute("noValue", "0"); $p.SetAttribute("flags", "3"); $p.SetAttribute("defValue", "1")
    $params += $p

    # Common globals after First Wave (same for all variants)
    $commonGlobals = @(
        @{name="Trig Type"; type="2"; min="0"; max="48"; no="255"; flags="2"; def="0"},
        @{name="Solo Track"; type="2"; min="0"; max=$numChannels.ToString(); no="255"; flags="2"; def=$numChannels.ToString()},
        @{name="Volume"; type="2"; min="0"; max="254"; no="255"; flags="2"; def="128"},
        @{name="Velocity"; type="2"; min="0"; max="254"; no="255"; flags="2"; def="128"},
        @{name="Human Vel"; type="2"; min="0"; max="128"; no="255"; flags="2"; def="0"},
        @{name="Pan"; type="2"; min="1"; max="255"; no="0"; flags="2"; def="128"},
        @{name="Human Pan"; type="2"; min="0"; max="128"; no="255"; flags="2"; def="0"},
        @{name="Tune"; type="2"; min="1"; max="255"; no="0"; flags="2"; def="128"},
        @{name="Human Tune"; type="2"; min="0"; max="128"; no="255"; flags="2"; def="0"},
        @{name="Len Unit"; type="2"; min="1"; max="20"; no="255"; flags="2"; def="6"},
        @{name="Shuffle Size"; type="2"; min="0"; max="254"; no="255"; flags="2"; def="0"},
        @{name="Shuffle Step"; type="2"; min="2"; max="32"; no="255"; flags="2"; def="2"},
        @{name="Shuffle Rnd"; type="2"; min="0"; max="254"; no="255"; flags="2"; def="0"},
        @{name="Shuffle Reset"; type="1"; min="-1"; max="-1"; no="255"; flags="0"; def="0"},
        @{name="Inertia"; type="3"; min="0"; max="1280"; no="65535"; flags="2"; def="0"}
    )

    foreach ($g in $commonGlobals) {
        $p = $Xml.CreateElement("Parameter"); $p.SetAttribute("scope", "global"); $p.SetAttribute("type", $g.type)
        $p.SetAttribute("name", $g.name); $p.SetAttribute("minValue", $g.min); $p.SetAttribute("maxValue", $g.max)
        $p.SetAttribute("noValue", $g.no); $p.SetAttribute("flags", $g.flags); $p.SetAttribute("defValue", $g.def)
        $params += $p
    }

    # v3 track params (26 params, no dividers)
    # Track Trigger type matches global trigger type: B variants = byte (2), S variants = switch (1)
    $trackTrigType = if ($isBVariant) { "2" } else { "1" }
    $trackParams = @(
        @{name="Trigger"; type=$trackTrigType; min="0"; max="254"; no="255"; flags="0"; def="0"},
        @{name="Wave No"; type="2"; min="0"; max="200"; no="255"; flags="3"; def="0"},
        @{name="Command 1"; type="2"; min="0"; max="254"; no="255"; flags="0"; def="0"},
        @{name="Argument 1"; type="2"; min="1"; max="255"; no="0"; flags="0"; def="0"},
        @{name="Command 2"; type="2"; min="0"; max="254"; no="255"; flags="0"; def="0"},
        @{name="Argument 2"; type="2"; min="1"; max="255"; no="0"; flags="0"; def="0"},
        @{name="Len Unit"; type="2"; min="0"; max="20"; no="255"; flags="2"; def="0"},
        @{name="Mute"; type="1"; min="-1"; max="-1"; no="255"; flags="2"; def="0"},
        @{name="Group"; type="2"; min="0"; max="15"; no="255"; flags="2"; def="0"},
        @{name="Start"; type="3"; min="0"; max="65534"; no="65535"; flags="2"; def="0"},
        @{name="Loop Start"; type="3"; min="0"; max="65534"; no="65535"; flags="2"; def="0"},
        @{name="End"; type="3"; min="1"; max="65535"; no="0"; flags="2"; def="65535"},
        @{name="Loop Mode"; type="2"; min="0"; max="3"; no="255"; flags="2"; def="0"},
        @{name="Note Cut"; type="2"; min="0"; max="254"; no="255"; flags="2"; def="0"},
        @{name="Loop Fit"; type="3"; min="0"; max="8192"; no="65535"; flags="2"; def="0"},
        @{name="Volume"; type="2"; min="0"; max="254"; no="255"; flags="2"; def="128"},
        @{name="Velocity"; type="2"; min="0"; max="254"; no="255"; flags="2"; def="128"},
        @{name="Human Vel"; type="2"; min="0"; max="128"; no="255"; flags="2"; def="0"},
        @{name="Vol Env"; type="2"; min="0"; max="64"; no="255"; flags="2"; def="0"},
        @{name="  Len"; type="2"; min="1"; max="255"; no="0"; flags="2"; def="16"},
        @{name="Pan"; type="2"; min="1"; max="255"; no="0"; flags="2"; def="128"},
        @{name="Human Pan"; type="2"; min="0"; max="128"; no="255"; flags="2"; def="0"},
        @{name="Pan Env"; type="2"; min="0"; max="64"; no="255"; flags="2"; def="0"},
        @{name="  Len"; type="2"; min="1"; max="255"; no="0"; flags="2"; def="16"},
        @{name="Tune"; type="2"; min="1"; max="255"; no="0"; flags="2"; def="128"},
        @{name="Human Tune"; type="2"; min="0"; max="128"; no="255"; flags="2"; def="0"}
    )

    foreach ($t in $trackParams) {
        $p = $Xml.CreateElement("Parameter"); $p.SetAttribute("scope", "track"); $p.SetAttribute("type", $t.type)
        $p.SetAttribute("name", $t.name); $p.SetAttribute("minValue", $t.min); $p.SetAttribute("maxValue", $t.max)
        $p.SetAttribute("noValue", $t.no); $p.SetAttribute("flags", $t.flags); $p.SetAttribute("defValue", $t.def)
        $params += $p
    }

    return $params
}

# ============================================================================
# Remap GlobalState bytes: v2 (061110) -> v3
# v2 order: FirstWave, (Div), Trigger0..N-1, (Div), TrigType...Inertia
# v3 order: Trigger0..N-1, FirstWave, TrigType...Inertia
# ============================================================================
function Convert-SGridGlobalState {
    param(
        [byte[]]$V2Bytes,
        [int]$NumChannels
    )

    # v2 layout: [0]=FirstWave(1b), [1]=Divider(1b), [2..1+N]=Triggers(N bytes), [2+N]=Divider(1b), [3+N..end]=rest
    $firstWave = $V2Bytes[0]
    # Divider at index 1 (skip)
    $triggers = $V2Bytes[2..(1 + $NumChannels)]
    # Divider at index (2 + $NumChannels) (skip)
    $restStart = 3 + $NumChannels
    $rest = $V2Bytes[$restStart..($V2Bytes.Length - 1)]

    # v3 layout: Triggers(N bytes), FirstWave(1b), rest
    $v3Bytes = [byte[]]::new($triggers.Length + 1 + $rest.Length)
    $pos = 0
    foreach ($t in $triggers) { $v3Bytes[$pos++] = $t }
    $v3Bytes[$pos++] = $firstWave
    foreach ($r in $rest) { $v3Bytes[$pos++] = $r }

    return $v3Bytes
}

# ============================================================================
# Remap GlobalState bytes: OLDER v2 (pre-061110) -> v3
# Older v2 global (22 bytes for S4):
#   FirstWave(1) Div(1) Trig0..N-1(N) Div(1) TrigType(1) Solo(1) Vol(1)
#   Vel(1) HumanVel(1) Pan(1) HumanPan(1) Tune(1) HumanTune(1) ShufSize(1)
#   ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
#   [NO LenUnit between HumanTune and ShufSize]
#
# v3 global (N+17 bytes):
#   Trig0..N-1(N) FirstWave(1) TrigType(1) Solo(1) Vol(1) Vel(1)
#   HumanVel(1) Pan(1) HumanPan(1) Tune(1) HumanTune(1) LenUnit(1)
#   ShufSize(1) ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
# ============================================================================
function Convert-SGridOlderGlobalState {
    param(
        [byte[]]$V2Bytes,
        [int]$NumChannels
    )

    $v3Size = $NumChannels + 17
    $v3Bytes = [byte[]]::new($v3Size)

    # v2 offsets
    $firstWave = $V2Bytes[0]
    # Triggers: v2[2..1+N] -> v3[0..N-1]
    for ($i = 0; $i -lt $NumChannels; $i++) {
        $v3Bytes[$i] = $V2Bytes[2 + $i]
    }
    # FirstWave -> v3[N]
    $v3Bytes[$NumChannels] = $firstWave

    # TrigType..HumanTune (9 bytes): v2 offset (3+N) -> v3 offset (N+1)
    $v2CommonStart = 3 + $NumChannels
    for ($b = 0; $b -lt 9; $b++) {
        $v3Bytes[$NumChannels + 1 + $b] = $V2Bytes[$v2CommonStart + $b]
    }
    # LenUnit: v3 offset (N+10) = default 0x06 (older has no LenUnit)
    $v3Bytes[$NumChannels + 10] = 0x06

    # ShufSize..ShufReset (4 bytes): v2 offset (3+N+9) -> v3 offset (N+11)
    for ($b = 0; $b -lt 4; $b++) {
        $v3Bytes[$NumChannels + 11 + $b] = $V2Bytes[$v2CommonStart + 9 + $b]
    }
    # Inertia (2 bytes): v2 offset (3+N+13) -> v3 offset (N+15)
    $v3Bytes[$NumChannels + 15] = $V2Bytes[$v2CommonStart + 13]
    $v3Bytes[$NumChannels + 16] = $V2Bytes[$v2CommonStart + 14]

    return $v3Bytes
}

# ============================================================================
# Remap TrackState bytes: v2 061110 (31 params, 35 bytes) -> v3 (26 params, 30 bytes)
# Remove divider bytes at offsets: 1, 10, 21, 27, 32
# ============================================================================
function Convert-SGridTrackState {
    param([byte[]]$V2Bytes)

    $dividerOffsets = @(1, 10, 21, 27, 32)
    $v3Bytes = [System.Collections.Generic.List[byte]]::new()
    for ($i = 0; $i -lt $V2Bytes.Length; $i++) {
        if ($i -notin $dividerOffsets) {
            $v3Bytes.Add($V2Bytes[$i])
        }
    }
    return [byte[]]$v3Bytes.ToArray()
}

# ============================================================================
# Remap TrackState bytes: OLDER v2 (15 bytes) -> v3 (26 params, 30 bytes)
# Older v2 track (15 bytes, no dividers, no trigger):
#   WaveNo(0) Cmd1(1) Arg1(2) Cmd2(3) Arg2(4) Subdiv(5) Mute(6)
#   Vol(7) Vel(8) HumanVel(9) Pan(10) HumanPan(11) Tune(12) HumanTune(13) Group(14)
#
# v3 track (30 bytes):
#   Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) LenUnit(6) Mute(7) Group(8)
#   Start(9,10) LoopStart(11,12) End(13,14) LoopMode(15) NoteCut(16) LoopFit(17,18)
#   Vol(19) Vel(20) HumanVel(21) VolEnv(22) VolEnvLen(23)
#   Pan(24) HumanPan(25) PanEnv(26) PanEnvLen(27) Tune(28) HumanTune(29)
# ============================================================================
function Convert-SGridOlderTrackState {
    param([byte[]]$V2Bytes)

    # Pre-fill with v3 default values (state uses defaults, not noValues)
    $v3Bytes = [byte[]]@(
        0xFF,       # Trigger (noVal - not a state param)
        0x00,       # Wave No (defVal=0)
        0xFF,       # Command 1 (noVal - not a state param)
        0x00,       # Argument 1 (defVal=0)
        0xFF,       # Command 2 (noVal - not a state param)
        0x00,       # Argument 2 (defVal=0)
        0x00,       # Len Unit (defVal=0)
        0x00,       # Mute (defVal=0)
        0x00,       # Group (defVal=0)
        0x00, 0x00, # Start (defVal=0)
        0x00, 0x00, # Loop Start (defVal=0)
        0xFF, 0xFF, # End (defVal=65535)
        0x00,       # Loop Mode (defVal=0)
        0x00,       # Note Cut (defVal=0)
        0x00, 0x00, # Loop Fit (defVal=0)
        0x80,       # Volume (defVal=128)
        0x80,       # Velocity (defVal=128)
        0x00,       # Human Vel (defVal=0)
        0x00,       # Vol Env (defVal=0)
        0x10,       # Vol Env Len (defVal=16)
        0x80,       # Pan (defVal=128)
        0x00,       # Human Pan (defVal=0)
        0x00,       # Pan Env (defVal=0)
        0x10,       # Pan Env Len (defVal=16)
        0x80,       # Tune (defVal=128)
        0x00        # Human Tune (defVal=0)
    )

    # Map older v2 offsets to v3 offsets
    $v3Bytes[1]  = $V2Bytes[0]   # WaveNo
    $v3Bytes[2]  = $V2Bytes[1]   # Cmd1
    $v3Bytes[3]  = $V2Bytes[2]   # Arg1
    $v3Bytes[4]  = $V2Bytes[3]   # Cmd2
    $v3Bytes[5]  = $V2Bytes[4]   # Arg2
    $v3Bytes[6]  = $V2Bytes[5]   # Subdiv -> LenUnit
    $v3Bytes[7]  = $V2Bytes[6]   # Mute
    $v3Bytes[8]  = $V2Bytes[14]  # Group
    $v3Bytes[19] = $V2Bytes[7]   # Volume
    $v3Bytes[20] = $V2Bytes[8]   # Velocity
    $v3Bytes[21] = $V2Bytes[9]   # HumanVel
    $v3Bytes[24] = $V2Bytes[10]  # Pan
    $v3Bytes[25] = $V2Bytes[11]  # HumanPan
    $v3Bytes[28] = $V2Bytes[12]  # Tune
    $v3Bytes[29] = $V2Bytes[13]  # HumanTune

    return $v3Bytes
}

# ============================================================================
# Remap GlobalState bytes: v1 SWITCH -> v3
# v1 switch global (N+11 bytes, NO dividers, NO TrigType):
#   FirstWave(1) Trig0..N-1(N) Solo(1) GlobalVol(1) GlobalPan(1) GlobalTune(1)
#   ShufSize(1) ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
#
# v3 global (N+17 bytes):
#   Trig0..N-1(N) FirstWave(1) TrigType(1) Solo(1) Vol(1) Vel(1)
#   HumanVel(1) Pan(1) HumanPan(1) Tune(1) HumanTune(1) LenUnit(1)
#   ShufSize(1) ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
# ============================================================================
function Convert-SGridV1SwitchGlobalState {
    param(
        [byte[]]$V1Bytes,
        [int]$NumChannels
    )

    $v3Size = $NumChannels + 17
    $v3Bytes = [byte[]]::new($v3Size)
    # Pre-fill with 0xFF for noValue defaults
    for ($i = 0; $i -lt $v3Size; $i++) { $v3Bytes[$i] = 0xFF }

    # v1 offsets: FirstWave(0), Triggers(1..N), Solo(N+1), GlobalVol(N+2),
    #   GlobalPan(N+3), GlobalTune(N+4), ShufSize(N+5)..ShufReset(N+8), Inertia(N+9,N+10)

    $firstWave = $V1Bytes[0]

    # Triggers -> v3[0..N-1]
    for ($i = 0; $i -lt $NumChannels; $i++) {
        $v3Bytes[$i] = $V1Bytes[1 + $i]
    }
    # FirstWave -> v3[N]
    $v3Bytes[$NumChannels] = $firstWave

    # TrigType -> v3[N+1] = default 0x00 (v1 switch has no TrigType)
    $v3Bytes[$NumChannels + 1] = 0x00
    # Solo -> v3[N+2]
    $v3Bytes[$NumChannels + 2] = $V1Bytes[$NumChannels + 1]
    # Vol -> v3[N+3] (from GlobalVol)
    $v3Bytes[$NumChannels + 3] = $V1Bytes[$NumChannels + 2]
    # Vel -> v3[N+4] = default 0x80 (128, v1 has no Velocity)
    $v3Bytes[$NumChannels + 4] = 0x80
    # HumanVel -> v3[N+5] = default 0x00 (v1 has no HumanVel)
    $v3Bytes[$NumChannels + 5] = 0x00
    # Pan -> v3[N+6] (from GlobalPan)
    $v3Bytes[$NumChannels + 6] = $V1Bytes[$NumChannels + 3]
    # HumanPan -> v3[N+7] = default 0x00
    $v3Bytes[$NumChannels + 7] = 0x00
    # Tune -> v3[N+8] (from GlobalTune)
    $v3Bytes[$NumChannels + 8] = $V1Bytes[$NumChannels + 4]
    # HumanTune -> v3[N+9] = default 0x00
    $v3Bytes[$NumChannels + 9] = 0x00
    # LenUnit -> v3[N+10] = default 0x06 (6)
    $v3Bytes[$NumChannels + 10] = 0x06
    # ShufSize..ShufReset (4 bytes) -> v3[N+11..N+14]
    for ($b = 0; $b -lt 4; $b++) {
        $v3Bytes[$NumChannels + 11 + $b] = $V1Bytes[$NumChannels + 5 + $b]
    }
    # Inertia (2 bytes) -> v3[N+15..N+16]
    $v3Bytes[$NumChannels + 15] = $V1Bytes[$NumChannels + 9]
    $v3Bytes[$NumChannels + 16] = $V1Bytes[$NumChannels + 10]

    return $v3Bytes
}

# ============================================================================
# Remap GlobalState bytes: v1 BYTE -> v3
# v1 byte global (N+13 bytes, HAS 2 dividers, HAS TrigType):
#   FirstWave(1) Div(1) Trig0..N-1(N) Div(1) TrigType(1) Solo(1) GlobalVol(1)
#   GlobalPan(1) GlobalTune(1) ShufSize(1) ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
#
# v3 global (N+17 bytes):
#   Trig0..N-1(N) FirstWave(1) TrigType(1) Solo(1) Vol(1) Vel(1)
#   HumanVel(1) Pan(1) HumanPan(1) Tune(1) HumanTune(1) LenUnit(1)
#   ShufSize(1) ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
# ============================================================================
function Convert-SGridV1ByteGlobalState {
    param(
        [byte[]]$V1Bytes,
        [int]$NumChannels
    )

    $v3Size = $NumChannels + 17
    $v3Bytes = [byte[]]::new($v3Size)
    # Pre-fill with 0xFF for noValue defaults
    for ($i = 0; $i -lt $v3Size; $i++) { $v3Bytes[$i] = 0xFF }

    # v1 byte offsets: FirstWave(0), Div(1), Triggers(2..1+N), Div(2+N),
    #   TrigType(3+N), Solo(4+N), GlobalVol(5+N), GlobalPan(6+N), GlobalTune(7+N)
    #   ShufSize(8+N)..ShufReset(11+N), Inertia(12+N, 13+N)

    $firstWave = $V1Bytes[0]

    # Triggers -> v3[0..N-1]
    for ($i = 0; $i -lt $NumChannels; $i++) {
        $v3Bytes[$i] = $V1Bytes[2 + $i]
    }
    # FirstWave -> v3[N]
    $v3Bytes[$NumChannels] = $firstWave

    # TrigType -> v3[N+1]
    $v3Bytes[$NumChannels + 1] = $V1Bytes[3 + $NumChannels]
    # Solo -> v3[N+2]
    $v3Bytes[$NumChannels + 2] = $V1Bytes[4 + $NumChannels]
    # Vol -> v3[N+3]
    $v3Bytes[$NumChannels + 3] = $V1Bytes[5 + $NumChannels]
    # Vel -> v3[N+4] = default 0x80 (128, v1 has no Velocity)
    $v3Bytes[$NumChannels + 4] = 0x80
    # HumanVel -> v3[N+5] = default 0x00
    $v3Bytes[$NumChannels + 5] = 0x00
    # Pan -> v3[N+6]
    $v3Bytes[$NumChannels + 6] = $V1Bytes[6 + $NumChannels]
    # HumanPan -> v3[N+7] = default 0x00
    $v3Bytes[$NumChannels + 7] = 0x00
    # Tune -> v3[N+8]
    $v3Bytes[$NumChannels + 8] = $V1Bytes[7 + $NumChannels]
    # HumanTune -> v3[N+9] = default 0x00
    $v3Bytes[$NumChannels + 9] = 0x00
    # LenUnit -> v3[N+10] = default 0x06 (6)
    $v3Bytes[$NumChannels + 10] = 0x06
    # ShufSize..ShufReset (4 bytes) -> v3[N+11..N+14]
    for ($b = 0; $b -lt 4; $b++) {
        $v3Bytes[$NumChannels + 11 + $b] = $V1Bytes[8 + $NumChannels + $b]
    }
    # Inertia (2 bytes) -> v3[N+15..N+16]
    $v3Bytes[$NumChannels + 15] = $V1Bytes[12 + $NumChannels]
    $v3Bytes[$NumChannels + 16] = $V1Bytes[13 + $NumChannels]

    return $v3Bytes
}

# ============================================================================
# Remap TrackState bytes: v1 (9 bytes) -> v3 (26 params, 30 bytes)
# v1 track (9 bytes, no dividers, no trigger):
#   WaveNo(0) Command(1) Argument(2) Subdiv(3) Mute(4)
#   Volume(5) Pan(6) Tune(7) AuxGroup(8)
#
# v3 track (30 bytes):
#   Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) LenUnit(6) Mute(7) Group(8)
#   Start(9,10) LoopStart(11,12) End(13,14) LoopMode(15) NoteCut(16) LoopFit(17,18)
#   Vol(19) Vel(20) HumanVel(21) VolEnv(22) VolEnvLen(23)
#   Pan(24) HumanPan(25) PanEnv(26) PanEnvLen(27) Tune(28) HumanTune(29)
# ============================================================================
function Convert-SGridV1TrackState {
    param([byte[]]$V1Bytes)

    # Pre-fill with v3 default values (state uses defaults, not noValues)
    $v3Bytes = [byte[]]@(
        0xFF,       # Trigger (noVal - not a state param)
        0x00,       # Wave No (defVal=0)
        0xFF,       # Command 1 (noVal - not a state param)
        0x00,       # Argument 1 (defVal=0)
        0xFF,       # Command 2 (noVal - not a state param)
        0x00,       # Argument 2 (defVal=0)
        0x00,       # Len Unit (defVal=0)
        0x00,       # Mute (defVal=0)
        0x00,       # Group (defVal=0)
        0x00, 0x00, # Start (defVal=0)
        0x00, 0x00, # Loop Start (defVal=0)
        0xFF, 0xFF, # End (defVal=65535)
        0x00,       # Loop Mode (defVal=0)
        0x00,       # Note Cut (defVal=0)
        0x00, 0x00, # Loop Fit (defVal=0)
        0x80,       # Volume (defVal=128)
        0x80,       # Velocity (defVal=128)
        0x00,       # Human Vel (defVal=0)
        0x00,       # Vol Env (defVal=0)
        0x10,       # Vol Env Len (defVal=16)
        0x80,       # Pan (defVal=128)
        0x00,       # Human Pan (defVal=0)
        0x00,       # Pan Env (defVal=0)
        0x10,       # Pan Env Len (defVal=16)
        0x80,       # Tune (defVal=128)
        0x00        # Human Tune (defVal=0)
    )

    # Map v1 offsets to v3 offsets
    $v3Bytes[1]  = $V1Bytes[0]   # WaveNo
    $v3Bytes[2]  = $V1Bytes[1]   # Command -> Cmd1
    $v3Bytes[3]  = $V1Bytes[2]   # Argument -> Arg1
    # v1 has no Cmd2/Arg2
    $v3Bytes[6]  = $V1Bytes[3]   # Subdiv -> LenUnit
    $v3Bytes[7]  = $V1Bytes[4]   # Mute
    $v3Bytes[8]  = $V1Bytes[8]   # AuxGroup -> Group
    $v3Bytes[19] = $V1Bytes[5]   # Volume
    # v1 has no Velocity, HumanVel
    $v3Bytes[24] = $V1Bytes[6]   # Pan
    # v1 has no HumanPan
    $v3Bytes[28] = $V1Bytes[7]   # Tune
    # v1 has no HumanTune

    return $v3Bytes
}

# ============================================================================
# Remap GlobalState bytes: v2 MID (050604) -> v3
# v2 mid global layout is same as older: has dividers, has TrigType
#   FirstWave(1) Div(1) Trig0..N-1(N) Div(1) TrigType(1) Solo(1) Vol(1)
#   Vel(1) HumanVel(1) Pan(1) HumanPan(1) Tune(1) HumanTune(1) ShufSize(1)
#   ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
#   [NO LenUnit between HumanTune and ShufSize - same as older]
# ============================================================================
# Uses Convert-SGridOlderGlobalState (identical global layout)

# ============================================================================
# Remap TrackState bytes: v2 MID (050604, 20 params, 23 bytes) -> v3 (30 bytes)
# v2 mid track (23 bytes, no dividers):
#   Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) Subdiv(6) Mute(7)
#   Offset(8,9:word) NoteCut(10) Vol(11) Vel(12) HumanVel(13) Pan(14)
#   HumanPan(15) Tune(16) HumanTune(17) LoopFit(18,19:word) LpFitMode(20,21:word) Group(22)
#
# v3 track (30 bytes):
#   Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) LenUnit(6) Mute(7) Group(8)
#   Start(9,10) LoopStart(11,12) End(13,14) LoopMode(15) NoteCut(16) LoopFit(17,18)
#   Vol(19) Vel(20) HumanVel(21) VolEnv(22) VolEnvLen(23)
#   Pan(24) HumanPan(25) PanEnv(26) PanEnvLen(27) Tune(28) HumanTune(29)
# ============================================================================
function Convert-SGridMidTrackState {
    param([byte[]]$V2Bytes)

    # Pre-fill with v3 default values (state uses defaults, not noValues)
    $v3Bytes = [byte[]]@(
        0xFF,       # Trigger (noVal - not a state param)
        0x00,       # Wave No (defVal=0)
        0xFF,       # Command 1 (noVal - not a state param)
        0x00,       # Argument 1 (defVal=0)
        0xFF,       # Command 2 (noVal - not a state param)
        0x00,       # Argument 2 (defVal=0)
        0x00,       # Len Unit (defVal=0)
        0x00,       # Mute (defVal=0)
        0x00,       # Group (defVal=0)
        0x00, 0x00, # Start (defVal=0)
        0x00, 0x00, # Loop Start (defVal=0)
        0xFF, 0xFF, # End (defVal=65535)
        0x00,       # Loop Mode (defVal=0)
        0x00,       # Note Cut (defVal=0)
        0x00, 0x00, # Loop Fit (defVal=0)
        0x80,       # Volume (defVal=128)
        0x80,       # Velocity (defVal=128)
        0x00,       # Human Vel (defVal=0)
        0x00,       # Vol Env (defVal=0)
        0x10,       # Vol Env Len (defVal=16)
        0x80,       # Pan (defVal=128)
        0x00,       # Human Pan (defVal=0)
        0x00,       # Pan Env (defVal=0)
        0x10,       # Pan Env Len (defVal=16)
        0x80,       # Tune (defVal=128)
        0x00        # Human Tune (defVal=0)
    )

    # Map v2-mid offsets to v3 offsets
    $v3Bytes[0]  = $V2Bytes[0]   # Trigger
    $v3Bytes[1]  = $V2Bytes[1]   # WaveNo
    $v3Bytes[2]  = $V2Bytes[2]   # Cmd1
    $v3Bytes[3]  = $V2Bytes[3]   # Arg1
    $v3Bytes[4]  = $V2Bytes[4]   # Cmd2
    $v3Bytes[5]  = $V2Bytes[5]   # Arg2
    $v3Bytes[6]  = $V2Bytes[6]   # Subdiv -> LenUnit
    $v3Bytes[7]  = $V2Bytes[7]   # Mute
    $v3Bytes[8]  = $V2Bytes[22]  # Group
    # Offset (word) -> Start (word)
    $v3Bytes[9]  = $V2Bytes[8]
    $v3Bytes[10] = $V2Bytes[9]
    $v3Bytes[16] = $V2Bytes[10]  # NoteCut
    # LoopFit (word)
    $v3Bytes[17] = $V2Bytes[18]
    $v3Bytes[18] = $V2Bytes[19]
    # LpFitMode (word in mid) -> DROPPED (not in v3)
    $v3Bytes[19] = $V2Bytes[11]  # Vol
    $v3Bytes[20] = $V2Bytes[12]  # Vel
    $v3Bytes[21] = $V2Bytes[13]  # HumanVel
    $v3Bytes[24] = $V2Bytes[14]  # Pan
    $v3Bytes[25] = $V2Bytes[15]  # HumanPan
    $v3Bytes[28] = $V2Bytes[16]  # Tune
    $v3Bytes[29] = $V2Bytes[17]  # HumanTune

    return $v3Bytes
}

# ============================================================================
# Remap TrackState bytes: v2 LATE (050716, 26 params, 28 bytes) -> v3 (30 bytes)
# v2 late track (28 bytes, no dividers):
#   Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) Subdiv(6) Mute(7)
#   Offset(8,9:word) NoteCut(10) Vol(11) Vel(12) HumanVel(13)
#   VolEnv(14) VolEnvLen(15) Pan(16) HumanPan(17) PanEnv(18) PanEnvLen(19)
#   Tune(20) HumanTune(21) TuneEnv(22) TuneEnvLen(23)
#   LoopFit(24,25:word) LpFitMode(26:byte) Group(27)
#
# v3 track (30 bytes):
#   Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) LenUnit(6) Mute(7) Group(8)
#   Start(9,10) LoopStart(11,12) End(13,14) LoopMode(15) NoteCut(16) LoopFit(17,18)
#   Vol(19) Vel(20) HumanVel(21) VolEnv(22) VolEnvLen(23)
#   Pan(24) HumanPan(25) PanEnv(26) PanEnvLen(27) Tune(28) HumanTune(29)
# ============================================================================
function Convert-SGridLateTrackState {
    param([byte[]]$V2Bytes)

    # Pre-fill with v3 default values (state uses defaults, not noValues)
    $v3Bytes = [byte[]]@(
        0xFF,       # Trigger (noVal - not a state param)
        0x00,       # Wave No (defVal=0)
        0xFF,       # Command 1 (noVal - not a state param)
        0x00,       # Argument 1 (defVal=0)
        0xFF,       # Command 2 (noVal - not a state param)
        0x00,       # Argument 2 (defVal=0)
        0x00,       # Len Unit (defVal=0)
        0x00,       # Mute (defVal=0)
        0x00,       # Group (defVal=0)
        0x00, 0x00, # Start (defVal=0)
        0x00, 0x00, # Loop Start (defVal=0)
        0xFF, 0xFF, # End (defVal=65535)
        0x00,       # Loop Mode (defVal=0)
        0x00,       # Note Cut (defVal=0)
        0x00, 0x00, # Loop Fit (defVal=0)
        0x80,       # Volume (defVal=128)
        0x80,       # Velocity (defVal=128)
        0x00,       # Human Vel (defVal=0)
        0x00,       # Vol Env (defVal=0)
        0x10,       # Vol Env Len (defVal=16)
        0x80,       # Pan (defVal=128)
        0x00,       # Human Pan (defVal=0)
        0x00,       # Pan Env (defVal=0)
        0x10,       # Pan Env Len (defVal=16)
        0x80,       # Tune (defVal=128)
        0x00        # Human Tune (defVal=0)
    )

    # Map v2-late offsets to v3 offsets
    $v3Bytes[0]  = $V2Bytes[0]   # Trigger
    $v3Bytes[1]  = $V2Bytes[1]   # WaveNo
    $v3Bytes[2]  = $V2Bytes[2]   # Cmd1
    $v3Bytes[3]  = $V2Bytes[3]   # Arg1
    $v3Bytes[4]  = $V2Bytes[4]   # Cmd2
    $v3Bytes[5]  = $V2Bytes[5]   # Arg2
    $v3Bytes[6]  = $V2Bytes[6]   # Subdiv -> LenUnit
    $v3Bytes[7]  = $V2Bytes[7]   # Mute
    $v3Bytes[8]  = $V2Bytes[27]  # Group
    # Offset (word) -> Start (word)
    $v3Bytes[9]  = $V2Bytes[8]
    $v3Bytes[10] = $V2Bytes[9]
    $v3Bytes[16] = $V2Bytes[10]  # NoteCut
    # LoopFit (word)
    $v3Bytes[17] = $V2Bytes[24]
    $v3Bytes[18] = $V2Bytes[25]
    # LpFitMode (byte in late) -> DROPPED
    # TuneEnv, TuneEnvLen -> DROPPED
    $v3Bytes[19] = $V2Bytes[11]  # Vol
    $v3Bytes[20] = $V2Bytes[12]  # Vel
    $v3Bytes[21] = $V2Bytes[13]  # HumanVel
    $v3Bytes[22] = $V2Bytes[14]  # VolEnv
    $v3Bytes[23] = $V2Bytes[15]  # VolEnvLen
    $v3Bytes[24] = $V2Bytes[16]  # Pan
    $v3Bytes[25] = $V2Bytes[17]  # HumanPan
    $v3Bytes[26] = $V2Bytes[18]  # PanEnv
    $v3Bytes[27] = $V2Bytes[19]  # PanEnvLen
    $v3Bytes[28] = $V2Bytes[20]  # Tune
    $v3Bytes[29] = $V2Bytes[21]  # HumanTune

    return $v3Bytes
}

# ============================================================================
# Build a default v3 data blob for SampleGrid 3 BETA 1
# The blob structure is:
#   [0]     version byte (0x07)
#   [1..N*5] MIDI key data (N*5 bytes, zeros = no assignments)
#   [N*5+1]  bKitMode (0x00 = no kit)
#   [N*5+2..N*5+1+N*11] peer data (N entries, 11 bytes each: 0x02 + 10 zeros)
#   [remainder] fixed 5210 bytes: group names, envelope defaults, etc.
# ============================================================================

# Fixed tail data (5210 bytes) - groups, envelopes, and output routing defaults
# Extracted from a fresh SampleGrid 3 BETA 1 instance (samplegrid_new_version.bmx)
# Stored in sgrid_v3_tail.b64 alongside this script

$script:SGridV3DataBlobTailBytes = $null
function Get-SGridV3DataBlobTailBytes {
    if (-not $script:SGridV3DataBlobTailBytes) {
        $tailFile = Join-Path (Split-Path $PSCommandPath -Parent) "sgrid_v3_tail.b64"
        $b64 = [System.IO.File]::ReadAllText($tailFile).Trim()
        $script:SGridV3DataBlobTailBytes = [Convert]::FromBase64String($b64)
    }
    return $script:SGridV3DataBlobTailBytes
}

function Build-SGridV3DefaultDataBlob {
    param([int]$NumChannels)

    $result = [System.Collections.Generic.List[byte]]::new()

    # Version byte
    $result.Add([byte]0x07)

    # MIDI key data: NumChannels * 5 bytes (all zeros = no assignments)
    for ($i = 0; $i -lt ($NumChannels * 5); $i++) {
        $result.Add([byte]0x00)
    }

    # bKitMode = 0 (no kit)
    $result.Add([byte]0x00)

    # Peer data: NumChannels entries, each 11 bytes (0x02 + 10 zeros)
    for ($ch = 0; $ch -lt $NumChannels; $ch++) {
        $result.Add([byte]0x02)
        for ($pad = 0; $pad -lt 10; $pad++) {
            $result.Add([byte]0x00)
        }
    }

    # Fixed tail: group names, envelope defaults, output routing (5210 bytes)
    $tailBytes = Get-SGridV3DataBlobTailBytes  # Get-SGridV3DataBlobTailBytes in SampleGridUpgrade.ps1
    foreach ($b in $tailBytes) {
        $result.Add($b)
    }

    return [byte[]]$result.ToArray()
}

# ============================================================================
# Convert machine data blob: v2 MDK format -> v3 format
# Differences:
#   - v2 has 1-byte MDK prefix (0x02) that v3 does not have
#   - v2 peer data: 2 bytes per entry; v3 peer data: 11 bytes per entry (2 + 9 zero pad)
#   - Everything else (version, MIDI, kit, groups, envelopes) is the same format
# ============================================================================
function Convert-SGridMachineData {
    param(
        [byte[]]$V2Data,
        [int]$NumTriggers
    )

    if ($V2Data.Length -lt 2) {
        Write-Log "    WARNING: Machine data too small ($($V2Data.Length) bytes), keeping as-is"
        return $V2Data
    }

    # Validate MDK prefix and version
    if ($V2Data[0] -ne 0x02) {
        Write-Log "    WARNING: Unexpected MDK prefix byte 0x$($V2Data[0].ToString('X2')) (expected 0x02), keeping as-is"
        return $V2Data
    }
    if ($V2Data[1] -notin @(0x07, 0x08)) {
        Write-Log "    WARNING: Unexpected SGrid version $($V2Data[1]) (expected 7 or 8), keeping as-is"
        return $V2Data
    }

    $v2Offset = 1  # skip MDK prefix byte; v3 starts with the version byte directly

    # Parse to find the peer data section boundaries
    $offset = 2  # start after version byte (relative to v2 data)

    # MIDI data: NumTriggers * 5 bytes (int nTrigKey=4 + byte nTrigKeyType=1)
    $midiSize = $NumTriggers * 5
    $offset += $midiSize

    # bKitMode: 1 byte
    if ($offset -ge $V2Data.Length) {
        Write-Log "    WARNING: Data ended before bKitMode, keeping as-is"
        return $V2Data
    }
    $bKitMode = $V2Data[$offset]
    $offset++

    if ($bKitMode -ne 0) {
        # nKitSaveMode: int (4 bytes)
        $kitSaveMode = [BitConverter]::ToInt32($V2Data, $offset)
        $offset += 4

        if ($kitSaveMode -eq 1 -or $kitSaveMode -eq 2) {
            # Embedded kit data - complex format, skip by finding end
            # This is harder to parse without knowing exact size
            # For now, log warning and keep data as-is
            Write-Log "    WARNING: Embedded kit data (mode $kitSaveMode) - data blob conversion not supported for embedded kits"
            return $V2Data
        } else {
            # Filename mode: read null-terminated string
            while ($offset -lt $V2Data.Length -and $V2Data[$offset] -ne 0) {
                $offset++
            }
            $offset++  # skip null terminator
        }
    }

    # $offset now points to peer data start
    $peerDataStart = $offset

    # v2 peer data: NumTriggers * NUM_PEERS(1) * 2 bytes each
    $v2PeerSize = $NumTriggers * 1 * 2
    $peerDataEnd = $peerDataStart + $v2PeerSize

    if ($peerDataEnd -gt $V2Data.Length) {
        Write-Log "    WARNING: Peer data extends beyond data blob, keeping as-is"
        return $V2Data
    }

    # v3 peer data: NumTriggers * NUM_PEERS(1) * 11 bytes each
    $v3PeerSize = $NumTriggers * 1 * 11

    # Build v3 data:
    # 1. Copy from v2[1] to v2[peerDataStart-1] (skip MDK prefix, keep version through kit data)
    # 2. Expand peer data entries from 2 bytes to 11 bytes (pad with 9 zeros)
    # 3. Copy remaining data (groups + envelopes) unchanged

    $result = [System.Collections.Generic.List[byte]]::new()

    # Part 1: version through pre-peer data (skip MDK prefix at byte 0)
    for ($i = $v2Offset; $i -lt $peerDataStart; $i++) {
        $result.Add($V2Data[$i])
    }

    # Part 2: expand peer data (2-byte entries -> 11-byte entries)
    for ($entry = 0; $entry -lt $NumTriggers; $entry++) {
        $entryOffset = $peerDataStart + ($entry * 2)
        # Copy original 2 bytes
        $result.Add($V2Data[$entryOffset])
        $result.Add($V2Data[$entryOffset + 1])
        # Pad with 9 zero bytes
        for ($pad = 0; $pad -lt 9; $pad++) {
            $result.Add([byte]0)
        }
    }

    # Part 3: copy remaining data (groups + envelopes)
    for ($i = $peerDataEnd; $i -lt $V2Data.Length; $i++) {
        $result.Add($V2Data[$i])
    }

    $v3Data = [byte[]]$result.ToArray()
    Write-Log "    Machine data blob: $($V2Data.Length) -> $($v3Data.Length) bytes (stripped MDK prefix, expanded peer data)"
    return $v3Data
}

# ============================================================================
# Remap PATT parameter data row: strip divider bytes from each row
# The PATT ParamData blob has rows of globalRowSize + (trackRowSize * numTracks) bytes
# v2 globalRowSize includes 2 divider bytes, v2 trackRowSize includes 5 divider bytes
# ============================================================================
function Convert-SGridPattData {
    param(
        [byte[]]$V2Data,
        [int]$V2GlobalRowSize,
        [int]$V2TrackRowSize,
        [int]$V3GlobalRowSize,
        [int]$V3TrackRowSize,
        [int]$NumTracks,
        [int]$NumRows,
        [int]$NumChannels
    )

    $v2RowSize = $V2GlobalRowSize + ($V2TrackRowSize * $NumTracks)
    $v3RowSize = $V3GlobalRowSize + ($V3TrackRowSize * $NumTracks)
    $v3Data = [byte[]]::new($v3RowSize * $NumRows)

    # Global divider offsets within the global portion of a row
    # v2 global: FirstWave(1b), Div(1b), Triggers(N*1b), Div(1b), rest...
    $globalDividerOffsets = @(1, (2 + $NumChannels))

    # Track divider offsets within each track's portion of a row (same as TrackState)
    $trackDividerOffsets = @(1, 10, 21, 27, 32)

    for ($row = 0; $row -lt $NumRows; $row++) {
        $v2RowStart = $row * $v2RowSize
        $v3RowStart = $row * $v3RowSize

        # Copy global params, skipping dividers
        # v2 global -> v3 global: reorder (FirstWave, Div, Triggers, Div, rest) -> (Triggers, FirstWave, rest)
        $firstWave = $V2Data[$v2RowStart]
        # Skip divider at offset 1
        $triggerBytes = @()
        for ($t = 0; $t -lt $NumChannels; $t++) {
            $triggerBytes += $V2Data[$v2RowStart + 2 + $t]
        }
        # Skip divider at offset (2 + NumChannels)
        $restStart = $v2RowStart + 3 + $NumChannels
        $restLen = $V2GlobalRowSize - 3 - $NumChannels

        # Write v3 global: Triggers, FirstWave, rest
        $v3Pos = $v3RowStart
        foreach ($tb in $triggerBytes) { $v3Data[$v3Pos++] = $tb }
        $v3Data[$v3Pos++] = $firstWave
        for ($r = 0; $r -lt $restLen; $r++) {
            $v3Data[$v3Pos++] = $V2Data[$restStart + $r]
        }

        # Copy track params, skipping dividers
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v2TrackStart = $v2RowStart + $V2GlobalRowSize + ($track * $V2TrackRowSize)
            $v3TrackStart = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)
            $v3TPos = $v3TrackStart
            for ($b = 0; $b -lt $V2TrackRowSize; $b++) {
                if ($b -notin $trackDividerOffsets) {
                    $v3Data[$v3TPos++] = $V2Data[$v2TrackStart + $b]
                }
            }
        }
    }

    # Sanitize: replace 0x00 with noValue for params where noValue != 0
    # v3 uninitialized pattern bytes are 0x00, but for params where 0 is a valid
    # (non-noValue) setting, v3 interprets them as real values (e.g. Solo Track 0
    # = "solo on track 0" instead of "no change"). Replace with noValue to fix.
    $v3Data = Set-SGridNoValues -V3Data $v3Data -V3GlobalRowSize $V3GlobalRowSize `
        -V3TrackRowSize $V3TrackRowSize -NumTracks $NumTracks -NumRows $NumRows `
        -NumChannels $NumChannels

    return $v3Data
}

# ============================================================================
# Build noValue byte map for v3 row layout, then replace 0x00 garbage with noValue
# Only replaces bytes where: value == 0x00 AND noValue != 0x00
# ============================================================================
function Set-SGridNoValues {
    param(
        [byte[]]$V3Data,
        [int]$V3GlobalRowSize,
        [int]$V3TrackRowSize,
        [int]$NumTracks,
        [int]$NumRows,
        [int]$NumChannels
    )

    # Build v3 global noValue byte array
    # v3 global byte order: Trigger0..TriggerN-1 (1b each, noVal=0xFF),
    #   FirstWave (1b, noVal=0x00),
    #   TrigType(1b,0xFF), Solo(1b,0xFF), Vol(1b,0xFF), Vel(1b,0xFF),
    #   HumanVel(1b,0xFF), Pan(1b,0x00), HumanPan(1b,0xFF), Tune(1b,0x00),
    #   HumanTune(1b,0xFF), LenUnit(1b,0xFF), ShuffleSize(1b,0xFF),
    #   ShuffleStep(1b,0xFF), ShuffleRnd(1b,0xFF), ShuffleReset(1b,0xFF),
    #   Inertia(2b, noVal=0xFF,0xFF)
    $globalNoVals = [System.Collections.Generic.List[byte]]::new()
    for ($i = 0; $i -lt $NumChannels; $i++) { $globalNoVals.Add([byte]0xFF) }  # Triggers
    $globalNoVals.Add([byte]0x00)  # First Wave (noVal=0)
    # TrigType, Solo, Vol, Vel, HumanVel
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }
    $globalNoVals.Add([byte]0x00)  # Pan (noVal=0)
    $globalNoVals.Add([byte]0xFF)  # HumanPan
    $globalNoVals.Add([byte]0x00)  # Tune (noVal=0)
    # HumanTune, LenUnit, ShuffleSize, ShuffleStep, ShuffleRnd, ShuffleReset
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }
    # Inertia (word, noVal=0xFFFF)
    $globalNoVals.Add([byte]0xFF); $globalNoVals.Add([byte]0xFF)

    # Build v3 track noValue byte array (26 params, 30 bytes, no dividers)
    # Trigger(1b,0xFF), WaveNo(1b,0xFF), Cmd1(1b,0xFF), Arg1(1b,0x00),
    # Cmd2(1b,0xFF), Arg2(1b,0x00), LenUnit(1b,0xFF), Mute(1b,0xFF),
    # Group(1b,0xFF), Start(2b,0xFF,0xFF), LoopStart(2b,0xFF,0xFF),
    # End(2b,0x00,0x00), LoopMode(1b,0xFF), NoteCut(1b,0xFF),
    # LoopFit(2b,0xFF,0xFF), Vol(1b,0xFF), Vel(1b,0xFF), HumanVel(1b,0xFF),
    # VolEnv(1b,0xFF), VolEnvLen(1b,0x00), Pan(1b,0x00), HumanPan(1b,0xFF),
    # PanEnv(1b,0xFF), PanEnvLen(1b,0x00), Tune(1b,0x00), HumanTune(1b,0xFF)
    $trackNoVals = [byte[]]@(
        0xFF,       # Trigger
        0xFF,       # Wave No
        0xFF,       # Command 1
        0x00,       # Argument 1 (noVal=0)
        0xFF,       # Command 2
        0x00,       # Argument 2 (noVal=0)
        0xFF,       # Len Unit
        0xFF,       # Mute
        0xFF,       # Group
        0xFF, 0xFF, # Start (word)
        0xFF, 0xFF, # Loop Start (word)
        0x00, 0x00, # End (word, noVal=0)
        0xFF,       # Loop Mode
        0xFF,       # Note Cut
        0xFF, 0xFF, # Loop Fit (word)
        0xFF,       # Volume
        0xFF,       # Velocity
        0xFF,       # Human Vel
        0xFF,       # Vol Env
        0x00,       # Vol Env Len (noVal=0)
        0x00,       # Pan (noVal=0)
        0xFF,       # Human Pan
        0xFF,       # Pan Env
        0x00,       # Pan Env Len (noVal=0)
        0x00,       # Tune (noVal=0)
        0xFF        # Human Tune
    )

    $v3RowSize = $V3GlobalRowSize + ($V3TrackRowSize * $NumTracks)
    $replaced = 0

    for ($row = 0; $row -lt $NumRows; $row++) {
        $rowStart = $row * $v3RowSize

        # Sanitize global params
        for ($b = 0; $b -lt $V3GlobalRowSize; $b++) {
            $pos = $rowStart + $b
            if ($V3Data[$pos] -eq 0x00 -and $globalNoVals[$b] -ne 0x00) {
                $V3Data[$pos] = $globalNoVals[$b]
                $replaced++
            }
        }

        # Sanitize track params
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $trackStart = $rowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)
            for ($b = 0; $b -lt $V3TrackRowSize; $b++) {
                $pos = $trackStart + $b
                if ($V3Data[$pos] -eq 0x00 -and $trackNoVals[$b] -ne 0x00) {
                    $V3Data[$pos] = $trackNoVals[$b]
                    $replaced++
                }
            }
        }
    }

    if ($replaced -gt 0) {
        Write-Log "    PATT sanitize: replaced $replaced garbage 0x00 bytes with noValue"
    }
    return $V3Data
}

# ============================================================================
# Convert PATT data for OLDER v2 versions (pre-061110, e.g. BETA 050422)
# Older v2 has a different param layout:
#   Globals (N+18 bytes): FirstWave(1) Div(1) Triggers(N) Div(1) TrigType Solo
#     Vol Vel HumanVel Pan HumanPan Tune HumanTune ShufSize ShufStep ShufRnd
#     ShufReset Inertia(2)  [NO LenUnit]
#   Tracks (15 bytes each, NO dividers, NO trigger):
#     WaveNo Cmd1 Arg1 Cmd2 Arg2 Subdiv Mute Vol Vel HumanVel Pan HumanPan
#     Tune HumanTune Group
# ============================================================================
function Convert-SGridOlderPattData {
    param(
        [byte[]]$V2Data,
        [int]$V2GlobalRowSize,
        [int]$V2TrackRowSize,
        [int]$V3GlobalRowSize,
        [int]$V3TrackRowSize,
        [int]$NumTracks,
        [int]$NumRows,
        [int]$NumChannels
    )

    $v2RowSize = $V2GlobalRowSize + ($V2TrackRowSize * $NumTracks)
    $v3RowSize = $V3GlobalRowSize + ($V3TrackRowSize * $NumTracks)
    $v3Data = [byte[]]::new($v3RowSize * $NumRows)

    # Pre-fill entire v3 data with noValue defaults so missing params are correct
    # v3 global noValues
    $globalNoVals = [System.Collections.Generic.List[byte]]::new()
    for ($i = 0; $i -lt $NumChannels; $i++) { $globalNoVals.Add([byte]0xFF) }  # Triggers
    $globalNoVals.Add([byte]0x00)  # First Wave
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }  # TrigType..HumanVel
    $globalNoVals.Add([byte]0x00)  # Pan
    $globalNoVals.Add([byte]0xFF)  # HumanPan
    $globalNoVals.Add([byte]0x00)  # Tune
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }  # HumanTune..ShufReset
    $globalNoVals.Add([byte]0xFF); $globalNoVals.Add([byte]0xFF)  # Inertia

    # v3 track noValues (30 bytes)
    $trackNoVals = [byte[]]@(
        0xFF,       # Trigger
        0xFF,       # Wave No
        0xFF,       # Command 1
        0x00,       # Argument 1
        0xFF,       # Command 2
        0x00,       # Argument 2
        0xFF,       # Len Unit
        0xFF,       # Mute
        0xFF,       # Group
        0xFF, 0xFF, # Start (word)
        0xFF, 0xFF, # Loop Start (word)
        0x00, 0x00, # End (word)
        0xFF,       # Loop Mode
        0xFF,       # Note Cut
        0xFF, 0xFF, # Loop Fit (word)
        0xFF,       # Volume
        0xFF,       # Velocity
        0xFF,       # Human Vel
        0xFF,       # Vol Env
        0x00,       # Vol Env Len
        0x00,       # Pan
        0xFF,       # Human Pan
        0xFF,       # Pan Env
        0x00,       # Pan Env Len
        0x00,       # Tune
        0xFF        # Human Tune
    )

    # Fill all rows with noValue defaults
    for ($row = 0; $row -lt $NumRows; $row++) {
        $v3RowStart = $row * $v3RowSize
        for ($b = 0; $b -lt $V3GlobalRowSize; $b++) {
            $v3Data[$v3RowStart + $b] = $globalNoVals[$b]
        }
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v3TrackStart = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)
            for ($b = 0; $b -lt $V3TrackRowSize; $b++) {
                $v3Data[$v3TrackStart + $b] = $trackNoVals[$b]
            }
        }
    }

    # Now copy v2 data into the correct v3 positions
    for ($row = 0; $row -lt $NumRows; $row++) {
        $v2RowStart = $row * $v2RowSize
        $v3RowStart = $row * $v3RowSize

        # --- Globals ---
        # v2 older: FirstWave(0) Div(1) Trig0..TrigN-1(2..1+N) Div(2+N) TrigType(3+N)
        #   Solo(4+N) Vol Vel HumanVel Pan HumanPan Tune HumanTune
        #   ShufSize ShufStep ShufRnd ShufReset Inertia(2b)
        #   [NO LenUnit between HumanTune and ShufSize]
        #
        # v3: Trig0..N-1(0..N-1) FirstWave(N) TrigType(N+1) Solo(N+2) Vol Vel HumanVel
        #   Pan HumanPan Tune HumanTune LenUnit ShufSize ShufStep ShufRnd ShufReset Inertia(2b)

        $firstWave = $V2Data[$v2RowStart]
        # Triggers -> v3 positions 0..N-1
        for ($t = 0; $t -lt $NumChannels; $t++) {
            $v3Data[$v3RowStart + $t] = $V2Data[$v2RowStart + 2 + $t]
        }
        # FirstWave -> v3 position N
        $v3Data[$v3RowStart + $NumChannels] = $firstWave

        # TrigType..HumanTune (10 bytes) from v2 offset (3+N) -> v3 offset (N+1)
        $v2CommonStart = $v2RowStart + 3 + $NumChannels
        $v3CommonStart = $v3RowStart + $NumChannels + 1
        for ($b = 0; $b -lt 10; $b++) {
            $v3Data[$v3CommonStart + $b] = $V2Data[$v2CommonStart + $b]
        }
        # v3 LenUnit at offset (N+11) stays as noValue (0xFF) - older version has no LenUnit
        # ShufSize..ShufReset (4 bytes) from v2 offset (3+N+10) -> v3 offset (N+12)
        for ($b = 0; $b -lt 4; $b++) {
            $v3Data[$v3RowStart + $NumChannels + 12 + $b] = $V2Data[$v2CommonStart + 10 + $b]
        }
        # Inertia (2 bytes) from v2 offset (3+N+14) -> v3 offset (N+16)
        # Wait, ShufReset is 1 byte, then Inertia(2b)
        # v2 common after divider: TrigType(1) Solo(1) Vol(1) Vel(1) HumanVel(1) Pan(1)
        #   HumanPan(1) Tune(1) HumanTune(1) ShufSize(1) ShufStep(1) ShufRnd(1) ShufReset(1)
        #   Inertia(2) = 13 + 2 = 15 bytes total
        # v3 after FirstWave: TrigType(1) Solo(1) Vol(1) Vel(1) HumanVel(1) Pan(1)
        #   HumanPan(1) Tune(1) HumanTune(1) LenUnit(1) ShufSize(1) ShufStep(1) ShufRnd(1)
        #   ShufReset(1) Inertia(2) = 14 + 2 = 16 bytes total
        # So: v2 bytes 0-8 (TrigType..HumanTune, 9 params) -> v3 bytes 0-8
        #     v3 byte 9 = LenUnit (noValue, already filled)
        #     v2 bytes 9-12 (ShufSize..ShufReset, 4 params) -> v3 bytes 10-13
        #     v2 bytes 13-14 (Inertia, 2 bytes) -> v3 bytes 14-15
        # Let me redo this properly:
        for ($b = 0; $b -lt 9; $b++) {  # TrigType..HumanTune
            $v3Data[$v3CommonStart + $b] = $V2Data[$v2CommonStart + $b]
        }
        # Skip v3 offset 9 (LenUnit = noValue)
        for ($b = 0; $b -lt 4; $b++) {  # ShufSize..ShufReset
            $v3Data[$v3CommonStart + 10 + $b] = $V2Data[$v2CommonStart + 9 + $b]
        }
        # Inertia (2 bytes)
        $v3Data[$v3CommonStart + 14] = $V2Data[$v2CommonStart + 13]
        $v3Data[$v3CommonStart + 15] = $V2Data[$v2CommonStart + 14]

        # --- Tracks ---
        # v2 older track (15 bytes, no dividers):
        #   WaveNo(0) Cmd1(1) Arg1(2) Cmd2(3) Arg2(4) Subdiv(5) Mute(6)
        #   Vol(7) Vel(8) HumanVel(9) Pan(10) HumanPan(11) Tune(12) HumanTune(13) Group(14)
        #
        # v3 track (30 bytes):
        #   Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) LenUnit(6) Mute(7) Group(8)
        #   Start(9,10) LoopStart(11,12) End(13,14) LoopMode(15) NoteCut(16) LoopFit(17,18)
        #   Vol(19) Vel(20) HumanVel(21) VolEnv(22) VolEnvLen(23)
        #   Pan(24) HumanPan(25) PanEnv(26) PanEnvLen(27) Tune(28) HumanTune(29)
        #
        # Mapping (v2 offset -> v3 offset):
        #   WaveNo(0)->1, Cmd1(1)->2, Arg1(2)->3, Cmd2(3)->4, Arg2(4)->5
        #   Subdiv(5)->6 (maps to LenUnit), Mute(6)->7
        #   Vol(7)->19, Vel(8)->20, HumanVel(9)->21
        #   Pan(10)->24, HumanPan(11)->25, Tune(12)->28, HumanTune(13)->29
        #   Group(14)->8
        # Missing in v3 (filled with noValue): Trigger, Start, LoopStart, End, LoopMode,
        #   NoteCut, LoopFit, VolEnv, VolEnvLen, PanEnv, PanEnvLen

        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v2t = $v2RowStart + $V2GlobalRowSize + ($track * $V2TrackRowSize)
            $v3t = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)

            $v3Data[$v3t + 1]  = $V2Data[$v2t + 0]   # WaveNo
            $v3Data[$v3t + 2]  = $V2Data[$v2t + 1]   # Cmd1
            $v3Data[$v3t + 3]  = $V2Data[$v2t + 2]   # Arg1
            $v3Data[$v3t + 4]  = $V2Data[$v2t + 3]   # Cmd2
            $v3Data[$v3t + 5]  = $V2Data[$v2t + 4]   # Arg2
            $v3Data[$v3t + 6]  = $V2Data[$v2t + 5]   # Subdiv -> LenUnit
            $v3Data[$v3t + 7]  = $V2Data[$v2t + 6]   # Mute
            $v3Data[$v3t + 8]  = $V2Data[$v2t + 14]  # Group
            $v3Data[$v3t + 19] = $V2Data[$v2t + 7]   # Volume
            $v3Data[$v3t + 20] = $V2Data[$v2t + 8]   # Velocity
            $v3Data[$v3t + 21] = $V2Data[$v2t + 9]   # HumanVel
            $v3Data[$v3t + 24] = $V2Data[$v2t + 10]  # Pan
            $v3Data[$v3t + 25] = $V2Data[$v2t + 11]  # HumanPan
            $v3Data[$v3t + 28] = $V2Data[$v2t + 12]  # Tune
            $v3Data[$v3t + 29] = $V2Data[$v2t + 13]  # HumanTune
        }
    }

    return $v3Data
}

# ============================================================================
# Convert PATT data for v1 SWITCH versions
# v1 switch global (N+11 bytes, NO dividers, NO TrigType):
#   FirstWave(1) Trig0..N-1(N) Solo(1) GlobalVol(1) GlobalPan(1) GlobalTune(1)
#   ShufSize(1) ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
# v1 track (9 bytes):
#   WaveNo(1) Command(1) Argument(1) Subdiv(1) Mute(1)
#   Volume(1) Pan(1) Tune(1) AuxGroup(1)
# ============================================================================
function Convert-SGridV1SwitchPattData {
    param(
        [byte[]]$V2Data,
        [int]$V2GlobalRowSize,
        [int]$V2TrackRowSize,
        [int]$V3GlobalRowSize,
        [int]$V3TrackRowSize,
        [int]$NumTracks,
        [int]$NumRows,
        [int]$NumChannels
    )

    # Alias for clarity (this is v1 data, but param named V2Data for dispatcher splatting)
    $V1Data = $V2Data
    $V1GlobalRowSize = $V2GlobalRowSize
    $V1TrackRowSize = $V2TrackRowSize

    $v1RowSize = $V1GlobalRowSize + ($V1TrackRowSize * $NumTracks)
    $v3RowSize = $V3GlobalRowSize + ($V3TrackRowSize * $NumTracks)
    $v3Data = [byte[]]::new($v3RowSize * $NumRows)

    # Pre-fill with noValue defaults (reuse same noValue arrays)
    $globalNoVals = [System.Collections.Generic.List[byte]]::new()
    for ($i = 0; $i -lt $NumChannels; $i++) { $globalNoVals.Add([byte]0xFF) }
    $globalNoVals.Add([byte]0x00)  # First Wave
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }  # TrigType..HumanVel
    $globalNoVals.Add([byte]0x00)  # Pan
    $globalNoVals.Add([byte]0xFF)  # HumanPan
    $globalNoVals.Add([byte]0x00)  # Tune
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }  # HumanTune..ShufReset
    $globalNoVals.Add([byte]0xFF); $globalNoVals.Add([byte]0xFF)  # Inertia

    $trackNoVals = [byte[]]@(
        0xFF, 0xFF, 0xFF, 0x00, 0xFF, 0x00, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0x00, 0x00, 0xFF
    )

    # Fill all rows with noValue defaults
    for ($row = 0; $row -lt $NumRows; $row++) {
        $v3RowStart = $row * $v3RowSize
        for ($b = 0; $b -lt $V3GlobalRowSize; $b++) { $v3Data[$v3RowStart + $b] = $globalNoVals[$b] }
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v3TrackStart = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)
            for ($b = 0; $b -lt $V3TrackRowSize; $b++) { $v3Data[$v3TrackStart + $b] = $trackNoVals[$b] }
        }
    }

    # Copy v1 data into v3 positions
    for ($row = 0; $row -lt $NumRows; $row++) {
        $v1RowStart = $row * $v1RowSize
        $v3RowStart = $row * $v3RowSize

        # --- Globals ---
        # v1 switch: FirstWave(0) Trig0..N-1(1..N) Solo(N+1) GlobalVol(N+2) GlobalPan(N+3)
        #   GlobalTune(N+4) ShufSize(N+5)..ShufReset(N+8) Inertia(N+9,N+10)
        # v3: Trig0..N-1(0..N-1) FirstWave(N) TrigType(N+1) Solo(N+2) Vol(N+3)
        #   Vel(N+4) HumanVel(N+5) Pan(N+6) HumanPan(N+7) Tune(N+8) HumanTune(N+9)
        #   LenUnit(N+10) ShufSize(N+11)..ShufReset(N+14) Inertia(N+15,N+16)

        # Triggers -> v3[0..N-1]
        for ($t = 0; $t -lt $NumChannels; $t++) {
            $v3Data[$v3RowStart + $t] = $V1Data[$v1RowStart + 1 + $t]
        }
        # FirstWave -> v3[N]
        $v3Data[$v3RowStart + $NumChannels] = $V1Data[$v1RowStart]
        # TrigType -> v3[N+1] stays noValue (0xFF)
        # Solo -> v3[N+2]
        $v3Data[$v3RowStart + $NumChannels + 2] = $V1Data[$v1RowStart + $NumChannels + 1]
        # Vol -> v3[N+3]
        $v3Data[$v3RowStart + $NumChannels + 3] = $V1Data[$v1RowStart + $NumChannels + 2]
        # Vel, HumanVel -> stay noValue
        # Pan -> v3[N+6]
        $v3Data[$v3RowStart + $NumChannels + 6] = $V1Data[$v1RowStart + $NumChannels + 3]
        # HumanPan -> stays noValue
        # Tune -> v3[N+8]
        $v3Data[$v3RowStart + $NumChannels + 8] = $V1Data[$v1RowStart + $NumChannels + 4]
        # HumanTune, LenUnit -> stay noValue
        # ShufSize..ShufReset (4 bytes)
        for ($b = 0; $b -lt 4; $b++) {
            $v3Data[$v3RowStart + $NumChannels + 11 + $b] = $V1Data[$v1RowStart + $NumChannels + 5 + $b]
        }
        # Inertia (2 bytes)
        $v3Data[$v3RowStart + $NumChannels + 15] = $V1Data[$v1RowStart + $NumChannels + 9]
        $v3Data[$v3RowStart + $NumChannels + 16] = $V1Data[$v1RowStart + $NumChannels + 10]

        # --- Tracks ---
        # v1: WaveNo(0) Command(1) Argument(2) Subdiv(3) Mute(4) Vol(5) Pan(6) Tune(7) AuxGroup(8)
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v1t = $v1RowStart + $V1GlobalRowSize + ($track * $V1TrackRowSize)
            $v3t = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)

            $v3Data[$v3t + 1]  = $V1Data[$v1t + 0]   # WaveNo
            $v3Data[$v3t + 2]  = $V1Data[$v1t + 1]   # Command -> Cmd1
            $v3Data[$v3t + 3]  = $V1Data[$v1t + 2]   # Argument -> Arg1
            $v3Data[$v3t + 6]  = $V1Data[$v1t + 3]   # Subdiv -> LenUnit
            $v3Data[$v3t + 7]  = $V1Data[$v1t + 4]   # Mute
            $v3Data[$v3t + 8]  = $V1Data[$v1t + 8]   # AuxGroup -> Group
            $v3Data[$v3t + 19] = $V1Data[$v1t + 5]   # Volume
            $v3Data[$v3t + 24] = $V1Data[$v1t + 6]   # Pan
            $v3Data[$v3t + 28] = $V1Data[$v1t + 7]   # Tune
        }
    }

    return $v3Data
}

# ============================================================================
# Convert PATT data for v1 BYTE versions
# v1 byte global (N+13 bytes, HAS 2 dividers, HAS TrigType):
#   FirstWave(1) Div(1) Trig0..N-1(N) Div(1) TrigType(1) Solo(1) GlobalVol(1)
#   GlobalPan(1) GlobalTune(1) ShufSize(1) ShufStep(1) ShufRnd(1) ShufReset(1) Inertia(2)
# v1 track is same as v1 switch (9 bytes, same mapping)
# ============================================================================
function Convert-SGridV1BytePattData {
    param(
        [byte[]]$V2Data,
        [int]$V2GlobalRowSize,
        [int]$V2TrackRowSize,
        [int]$V3GlobalRowSize,
        [int]$V3TrackRowSize,
        [int]$NumTracks,
        [int]$NumRows,
        [int]$NumChannels
    )

    # Alias for clarity
    $V1Data = $V2Data
    $V1GlobalRowSize = $V2GlobalRowSize
    $V1TrackRowSize = $V2TrackRowSize

    $v1RowSize = $V1GlobalRowSize + ($V1TrackRowSize * $NumTracks)
    $v3RowSize = $V3GlobalRowSize + ($V3TrackRowSize * $NumTracks)
    $v3Data = [byte[]]::new($v3RowSize * $NumRows)

    # Pre-fill with noValue defaults
    $globalNoVals = [System.Collections.Generic.List[byte]]::new()
    for ($i = 0; $i -lt $NumChannels; $i++) { $globalNoVals.Add([byte]0xFF) }
    $globalNoVals.Add([byte]0x00)  # First Wave
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }
    $globalNoVals.Add([byte]0x00)  # Pan
    $globalNoVals.Add([byte]0xFF)  # HumanPan
    $globalNoVals.Add([byte]0x00)  # Tune
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }
    $globalNoVals.Add([byte]0xFF); $globalNoVals.Add([byte]0xFF)

    $trackNoVals = [byte[]]@(
        0xFF, 0xFF, 0xFF, 0x00, 0xFF, 0x00, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0x00, 0x00, 0xFF
    )

    # Fill all rows with noValue defaults
    for ($row = 0; $row -lt $NumRows; $row++) {
        $v3RowStart = $row * $v3RowSize
        for ($b = 0; $b -lt $V3GlobalRowSize; $b++) { $v3Data[$v3RowStart + $b] = $globalNoVals[$b] }
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v3TrackStart = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)
            for ($b = 0; $b -lt $V3TrackRowSize; $b++) { $v3Data[$v3TrackStart + $b] = $trackNoVals[$b] }
        }
    }

    # Copy v1 data into v3 positions
    for ($row = 0; $row -lt $NumRows; $row++) {
        $v1RowStart = $row * $v1RowSize
        $v3RowStart = $row * $v3RowSize

        # --- Globals ---
        # v1 byte: FirstWave(0) Div(1) Trig0..N-1(2..1+N) Div(2+N)
        #   TrigType(3+N) Solo(4+N) GlobalVol(5+N) GlobalPan(6+N) GlobalTune(7+N)
        #   ShufSize(8+N)..ShufReset(11+N) Inertia(12+N,13+N)
        # v3: same as Convert-SGridOlderGlobalState mapping, but fewer params after TrigType

        # Triggers -> v3[0..N-1]
        for ($t = 0; $t -lt $NumChannels; $t++) {
            $v3Data[$v3RowStart + $t] = $V1Data[$v1RowStart + 2 + $t]
        }
        # FirstWave -> v3[N]
        $v3Data[$v3RowStart + $NumChannels] = $V1Data[$v1RowStart]
        # TrigType -> v3[N+1]
        $v2CommonStart = $v1RowStart + 3 + $NumChannels
        $v3Data[$v3RowStart + $NumChannels + 1] = $V1Data[$v2CommonStart]
        # Solo -> v3[N+2]
        $v3Data[$v3RowStart + $NumChannels + 2] = $V1Data[$v2CommonStart + 1]
        # Vol -> v3[N+3]
        $v3Data[$v3RowStart + $NumChannels + 3] = $V1Data[$v2CommonStart + 2]
        # Vel, HumanVel -> stay noValue
        # Pan -> v3[N+6]
        $v3Data[$v3RowStart + $NumChannels + 6] = $V1Data[$v2CommonStart + 3]
        # HumanPan -> stays noValue
        # Tune -> v3[N+8]
        $v3Data[$v3RowStart + $NumChannels + 8] = $V1Data[$v2CommonStart + 4]
        # HumanTune, LenUnit -> stay noValue
        # ShufSize..ShufReset (4 bytes)
        for ($b = 0; $b -lt 4; $b++) {
            $v3Data[$v3RowStart + $NumChannels + 11 + $b] = $V1Data[$v2CommonStart + 5 + $b]
        }
        # Inertia (2 bytes)
        $v3Data[$v3RowStart + $NumChannels + 15] = $V1Data[$v2CommonStart + 9]
        $v3Data[$v3RowStart + $NumChannels + 16] = $V1Data[$v2CommonStart + 10]

        # --- Tracks (same as v1 switch) ---
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v1t = $v1RowStart + $V1GlobalRowSize + ($track * $V1TrackRowSize)
            $v3t = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)

            $v3Data[$v3t + 1]  = $V1Data[$v1t + 0]   # WaveNo
            $v3Data[$v3t + 2]  = $V1Data[$v1t + 1]   # Command -> Cmd1
            $v3Data[$v3t + 3]  = $V1Data[$v1t + 2]   # Argument -> Arg1
            $v3Data[$v3t + 6]  = $V1Data[$v1t + 3]   # Subdiv -> LenUnit
            $v3Data[$v3t + 7]  = $V1Data[$v1t + 4]   # Mute
            $v3Data[$v3t + 8]  = $V1Data[$v1t + 8]   # AuxGroup -> Group
            $v3Data[$v3t + 19] = $V1Data[$v1t + 5]   # Volume
            $v3Data[$v3t + 24] = $V1Data[$v1t + 6]   # Pan
            $v3Data[$v3t + 28] = $V1Data[$v1t + 7]   # Tune
        }
    }

    return $v3Data
}

# ============================================================================
# Convert PATT data for v2 MID (050604, 20 track params, 23 bytes/track)
# Global layout is same as OLDER v2 (has dividers, TrigType, no LenUnit)
# Track: Trigger(1) WaveNo(1) Cmd1(1) Arg1(1) Cmd2(1) Arg2(1) Subdiv(1) Mute(1)
#   Offset(2:word) NoteCut(1) Vol(1) Vel(1) HumanVel(1) Pan(1) HumanPan(1) Tune(1)
#   HumanTune(1) LoopFit(2:word) LpFitMode(2:word) Group(1) = 23 bytes
# ============================================================================
function Convert-SGridMidPattData {
    param(
        [byte[]]$V2Data,
        [int]$V2GlobalRowSize,
        [int]$V2TrackRowSize,
        [int]$V3GlobalRowSize,
        [int]$V3TrackRowSize,
        [int]$NumTracks,
        [int]$NumRows,
        [int]$NumChannels
    )

    $v2RowSize = $V2GlobalRowSize + ($V2TrackRowSize * $NumTracks)
    $v3RowSize = $V3GlobalRowSize + ($V3TrackRowSize * $NumTracks)
    $v3Data = [byte[]]::new($v3RowSize * $NumRows)

    # Pre-fill with noValue defaults
    $globalNoVals = [System.Collections.Generic.List[byte]]::new()
    for ($i = 0; $i -lt $NumChannels; $i++) { $globalNoVals.Add([byte]0xFF) }
    $globalNoVals.Add([byte]0x00)  # First Wave
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }
    $globalNoVals.Add([byte]0x00)  # Pan
    $globalNoVals.Add([byte]0xFF)  # HumanPan
    $globalNoVals.Add([byte]0x00)  # Tune
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }
    $globalNoVals.Add([byte]0xFF); $globalNoVals.Add([byte]0xFF)

    $trackNoVals = [byte[]]@(
        0xFF, 0xFF, 0xFF, 0x00, 0xFF, 0x00, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0x00, 0x00, 0xFF
    )

    for ($row = 0; $row -lt $NumRows; $row++) {
        $v3RowStart = $row * $v3RowSize
        for ($b = 0; $b -lt $V3GlobalRowSize; $b++) { $v3Data[$v3RowStart + $b] = $globalNoVals[$b] }
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v3TrackStart = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)
            for ($b = 0; $b -lt $V3TrackRowSize; $b++) { $v3Data[$v3TrackStart + $b] = $trackNoVals[$b] }
        }
    }

    for ($row = 0; $row -lt $NumRows; $row++) {
        $v2RowStart = $row * $v2RowSize
        $v3RowStart = $row * $v3RowSize

        # --- Globals (same as older v2 - has dividers, has TrigType, no LenUnit) ---
        $v3Data[$v3RowStart + $NumChannels] = $V2Data[$v2RowStart]  # FirstWave
        for ($t = 0; $t -lt $NumChannels; $t++) {
            $v3Data[$v3RowStart + $t] = $V2Data[$v2RowStart + 2 + $t]  # Triggers
        }
        $v2CommonStart = $v2RowStart + 3 + $NumChannels
        $v3CommonStart = $v3RowStart + $NumChannels + 1
        for ($b = 0; $b -lt 9; $b++) {  # TrigType..HumanTune
            $v3Data[$v3CommonStart + $b] = $V2Data[$v2CommonStart + $b]
        }
        # Skip LenUnit (stays noValue)
        for ($b = 0; $b -lt 4; $b++) {  # ShufSize..ShufReset
            $v3Data[$v3CommonStart + 10 + $b] = $V2Data[$v2CommonStart + 9 + $b]
        }
        # Inertia (2 bytes)
        $v3Data[$v3CommonStart + 14] = $V2Data[$v2CommonStart + 13]
        $v3Data[$v3CommonStart + 15] = $V2Data[$v2CommonStart + 14]

        # --- Tracks ---
        # v2 mid: Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) Subdiv(6) Mute(7)
        #   Offset(8,9) NoteCut(10) Vol(11) Vel(12) HumanVel(13) Pan(14)
        #   HumanPan(15) Tune(16) HumanTune(17) LoopFit(18,19) LpFitMode(20,21) Group(22)
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v2t = $v2RowStart + $V2GlobalRowSize + ($track * $V2TrackRowSize)
            $v3t = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)

            $v3Data[$v3t + 0]  = $V2Data[$v2t + 0]   # Trigger
            $v3Data[$v3t + 1]  = $V2Data[$v2t + 1]   # WaveNo
            $v3Data[$v3t + 2]  = $V2Data[$v2t + 2]   # Cmd1
            $v3Data[$v3t + 3]  = $V2Data[$v2t + 3]   # Arg1
            $v3Data[$v3t + 4]  = $V2Data[$v2t + 4]   # Cmd2
            $v3Data[$v3t + 5]  = $V2Data[$v2t + 5]   # Arg2
            $v3Data[$v3t + 6]  = $V2Data[$v2t + 6]   # Subdiv -> LenUnit
            $v3Data[$v3t + 7]  = $V2Data[$v2t + 7]   # Mute
            $v3Data[$v3t + 8]  = $V2Data[$v2t + 22]  # Group
            # Offset (word) -> Start (word)
            $v3Data[$v3t + 9]  = $V2Data[$v2t + 8]
            $v3Data[$v3t + 10] = $V2Data[$v2t + 9]
            $v3Data[$v3t + 16] = $V2Data[$v2t + 10]  # NoteCut
            # LoopFit (word)
            $v3Data[$v3t + 17] = $V2Data[$v2t + 18]
            $v3Data[$v3t + 18] = $V2Data[$v2t + 19]
            # LpFitMode (word) -> DROPPED
            $v3Data[$v3t + 19] = $V2Data[$v2t + 11]  # Vol
            $v3Data[$v3t + 20] = $V2Data[$v2t + 12]  # Vel
            $v3Data[$v3t + 21] = $V2Data[$v2t + 13]  # HumanVel
            $v3Data[$v3t + 24] = $V2Data[$v2t + 14]  # Pan
            $v3Data[$v3t + 25] = $V2Data[$v2t + 15]  # HumanPan
            $v3Data[$v3t + 28] = $V2Data[$v2t + 16]  # Tune
            $v3Data[$v3t + 29] = $V2Data[$v2t + 17]  # HumanTune
        }
    }

    return $v3Data
}

# ============================================================================
# Convert PATT data for v2 LATE (050716, 26 track params, 28 bytes/track)
# Global layout same as OLDER v2 / MID
# Track: Trigger(1) WaveNo(1) Cmd1(1) Arg1(1) Cmd2(1) Arg2(1) Subdiv(1) Mute(1)
#   Offset(2:word) NoteCut(1) Vol(1) Vel(1) HumanVel(1) VolEnv(1) VolEnvLen(1)
#   Pan(1) HumanPan(1) PanEnv(1) PanEnvLen(1) Tune(1) HumanTune(1)
#   TuneEnv(1) TuneEnvLen(1) LoopFit(2:word) LpFitMode(1:byte) Group(1) = 28 bytes
# ============================================================================
function Convert-SGridLatePattData {
    param(
        [byte[]]$V2Data,
        [int]$V2GlobalRowSize,
        [int]$V2TrackRowSize,
        [int]$V3GlobalRowSize,
        [int]$V3TrackRowSize,
        [int]$NumTracks,
        [int]$NumRows,
        [int]$NumChannels
    )

    $v2RowSize = $V2GlobalRowSize + ($V2TrackRowSize * $NumTracks)
    $v3RowSize = $V3GlobalRowSize + ($V3TrackRowSize * $NumTracks)
    $v3Data = [byte[]]::new($v3RowSize * $NumRows)

    # Pre-fill with noValue defaults
    $globalNoVals = [System.Collections.Generic.List[byte]]::new()
    for ($i = 0; $i -lt $NumChannels; $i++) { $globalNoVals.Add([byte]0xFF) }
    $globalNoVals.Add([byte]0x00)
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }
    $globalNoVals.Add([byte]0x00)
    $globalNoVals.Add([byte]0xFF)
    $globalNoVals.Add([byte]0x00)
    foreach ($nv in @(0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF)) { $globalNoVals.Add([byte]$nv) }
    $globalNoVals.Add([byte]0xFF); $globalNoVals.Add([byte]0xFF)

    $trackNoVals = [byte[]]@(
        0xFF, 0xFF, 0xFF, 0x00, 0xFF, 0x00, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0x00, 0x00, 0xFF
    )

    for ($row = 0; $row -lt $NumRows; $row++) {
        $v3RowStart = $row * $v3RowSize
        for ($b = 0; $b -lt $V3GlobalRowSize; $b++) { $v3Data[$v3RowStart + $b] = $globalNoVals[$b] }
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v3TrackStart = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)
            for ($b = 0; $b -lt $V3TrackRowSize; $b++) { $v3Data[$v3TrackStart + $b] = $trackNoVals[$b] }
        }
    }

    for ($row = 0; $row -lt $NumRows; $row++) {
        $v2RowStart = $row * $v2RowSize
        $v3RowStart = $row * $v3RowSize

        # --- Globals (same as older v2 / mid) ---
        $v3Data[$v3RowStart + $NumChannels] = $V2Data[$v2RowStart]  # FirstWave
        for ($t = 0; $t -lt $NumChannels; $t++) {
            $v3Data[$v3RowStart + $t] = $V2Data[$v2RowStart + 2 + $t]
        }
        $v2CommonStart = $v2RowStart + 3 + $NumChannels
        $v3CommonStart = $v3RowStart + $NumChannels + 1
        for ($b = 0; $b -lt 9; $b++) {
            $v3Data[$v3CommonStart + $b] = $V2Data[$v2CommonStart + $b]
        }
        for ($b = 0; $b -lt 4; $b++) {
            $v3Data[$v3CommonStart + 10 + $b] = $V2Data[$v2CommonStart + 9 + $b]
        }
        $v3Data[$v3CommonStart + 14] = $V2Data[$v2CommonStart + 13]
        $v3Data[$v3CommonStart + 15] = $V2Data[$v2CommonStart + 14]

        # --- Tracks ---
        # v2 late: Trigger(0) WaveNo(1) Cmd1(2) Arg1(3) Cmd2(4) Arg2(5) Subdiv(6) Mute(7)
        #   Offset(8,9) NoteCut(10) Vol(11) Vel(12) HumanVel(13)
        #   VolEnv(14) VolEnvLen(15) Pan(16) HumanPan(17) PanEnv(18) PanEnvLen(19)
        #   Tune(20) HumanTune(21) TuneEnv(22) TuneEnvLen(23)
        #   LoopFit(24,25) LpFitMode(26) Group(27)
        for ($track = 0; $track -lt $NumTracks; $track++) {
            $v2t = $v2RowStart + $V2GlobalRowSize + ($track * $V2TrackRowSize)
            $v3t = $v3RowStart + $V3GlobalRowSize + ($track * $V3TrackRowSize)

            $v3Data[$v3t + 0]  = $V2Data[$v2t + 0]   # Trigger
            $v3Data[$v3t + 1]  = $V2Data[$v2t + 1]   # WaveNo
            $v3Data[$v3t + 2]  = $V2Data[$v2t + 2]   # Cmd1
            $v3Data[$v3t + 3]  = $V2Data[$v2t + 3]   # Arg1
            $v3Data[$v3t + 4]  = $V2Data[$v2t + 4]   # Cmd2
            $v3Data[$v3t + 5]  = $V2Data[$v2t + 5]   # Arg2
            $v3Data[$v3t + 6]  = $V2Data[$v2t + 6]   # Subdiv -> LenUnit
            $v3Data[$v3t + 7]  = $V2Data[$v2t + 7]   # Mute
            $v3Data[$v3t + 8]  = $V2Data[$v2t + 27]  # Group
            # Offset (word) -> Start (word)
            $v3Data[$v3t + 9]  = $V2Data[$v2t + 8]
            $v3Data[$v3t + 10] = $V2Data[$v2t + 9]
            $v3Data[$v3t + 16] = $V2Data[$v2t + 10]  # NoteCut
            # LoopFit (word)
            $v3Data[$v3t + 17] = $V2Data[$v2t + 24]
            $v3Data[$v3t + 18] = $V2Data[$v2t + 25]
            # LpFitMode -> DROPPED
            # TuneEnv, TuneEnvLen -> DROPPED
            $v3Data[$v3t + 19] = $V2Data[$v2t + 11]  # Vol
            $v3Data[$v3t + 20] = $V2Data[$v2t + 12]  # Vel
            $v3Data[$v3t + 21] = $V2Data[$v2t + 13]  # HumanVel
            $v3Data[$v3t + 22] = $V2Data[$v2t + 14]  # VolEnv
            $v3Data[$v3t + 23] = $V2Data[$v2t + 15]  # VolEnvLen
            $v3Data[$v3t + 24] = $V2Data[$v2t + 16]  # Pan
            $v3Data[$v3t + 25] = $V2Data[$v2t + 17]  # HumanPan
            $v3Data[$v3t + 26] = $V2Data[$v2t + 18]  # PanEnv
            $v3Data[$v3t + 27] = $V2Data[$v2t + 19]  # PanEnvLen
            $v3Data[$v3t + 28] = $V2Data[$v2t + 20]  # Tune
            $v3Data[$v3t + 29] = $V2Data[$v2t + 21]  # HumanTune
        }
    }

    return $v3Data
}

# ============================================================================
# Build flat paramIndex mapping from v2 to v3 for Pattern XP data conversion
# Pattern XP uses a flat index across all params (globals first, then tracks)
# Returns a hashtable: v2FlatIndex -> v3FlatIndex (or -1 for dividers to skip)
# ============================================================================
function Build-SGridParamIndexMap {
    param([int]$NumChannels)

    $map = @{}

    # --- Global params ---
    # v2 global layout (flat indices 0..numV2Global-1):
    #   0: FirstWave
    #   1: Divider
    #   2..(1+N): Triggers 0..(N-1)
    #   (2+N): Divider
    #   (3+N)..end: TrigType, Solo, Vol, Vel, HumanVel, Pan, HumanPan,
    #              Tune, HumanTune, LenUnit, ShufSize, ShufStep, ShufRnd,
    #              ShufReset, Inertia
    #
    # v3 global layout (flat indices 0..numV3Global-1):
    #   0..(N-1): Triggers 0..(N-1)
    #   N: FirstWave
    #   (N+1)..end: TrigType, Solo, Vol, Vel, HumanVel, Pan, HumanPan,
    #              Tune, HumanTune, LenUnit, ShufSize, ShufStep, ShufRnd,
    #              ShufReset, Inertia

    $numV2Global = $NumChannels + 2 + 15  # FirstWave + 2 dividers + N triggers + 15 common params
    # Actually: 1(FirstWave) + 1(Div) + N(Triggers) + 1(Div) + 15(TrigType..Inertia) = N + 18
    # Wait, let me count v2 common params after 2nd divider:
    # TrigType, Solo, Vol, Vel, HumanVel, Pan, HumanPan, Tune, HumanTune,
    # LenUnit, ShufSize, ShufStep, ShufRnd, ShufReset, Inertia = 15 params
    # Total v2 global = 1 + 1 + N + 1 + 15 = N + 18

    # v2 idx 0 = FirstWave -> v3 idx N
    $map[0] = $NumChannels

    # v2 idx 1 = Divider -> skip
    $map[1] = -1

    # v2 idx 2..(1+N) = Triggers -> v3 idx 0..(N-1)
    for ($i = 0; $i -lt $NumChannels; $i++) {
        $map[2 + $i] = $i
    }

    # v2 idx (2+N) = Divider -> skip
    $map[2 + $NumChannels] = -1

    # v2 idx (3+N)..end = common params -> v3 idx (N+1)..end
    # 15 common params
    for ($i = 0; $i -lt 15; $i++) {
        $map[3 + $NumChannels + $i] = $NumChannels + 1 + $i
    }

    $numV2Global = $NumChannels + 18
    $numV3Global = $NumChannels + 16  # N triggers + FirstWave + 15 common

    # --- Track params ---
    # v2 track layout (flat indices start at numV2Global):
    #   0: Trigger, 1: Divider, 2: WaveNo, 3: Cmd1, 4: Arg1, 5: Cmd2,
    #   6: Arg2, 7: LenUnit, 8: Mute, 9: Group,
    #   10: Divider, 11: Start, 12: LoopStart, 13: End, 14: LoopMode,
    #   15: NoteCut, 16: LoopFit,
    #   17: Divider, 18: Vol, 19: Vel, 20: HumanVel, 21: VolEnv, 22: VolEnvLen,
    #   23: Divider, 24: Pan, 25: HumanPan, 26: PanEnv, 27: PanEnvLen,
    #   28: Divider, 29: Tune, 30: HumanTune
    #   = 31 track params
    #
    # v3 track layout (flat indices start at numV3Global):
    #   0: Trigger, 1: WaveNo, 2: Cmd1, 3: Arg1, 4: Cmd2, 5: Arg2,
    #   6: LenUnit, 7: Mute, 8: Group,
    #   9: Start, 10: LoopStart, 11: End, 12: LoopMode, 13: NoteCut, 14: LoopFit,
    #   15: Vol, 16: Vel, 17: HumanVel, 18: VolEnv, 19: VolEnvLen,
    #   20: Pan, 21: HumanPan, 22: PanEnv, 23: PanEnvLen,
    #   24: Tune, 25: HumanTune
    #   = 26 track params

    $v2TrackDividers = @(1, 10, 17, 23, 28)  # 0-based within track params
    $v3TrackIdx = 0
    for ($v2t = 0; $v2t -lt 31; $v2t++) {
        $v2Flat = $numV2Global + $v2t
        if ($v2t -in $v2TrackDividers) {
            $map[$v2Flat] = -1  # skip divider
        } else {
            $map[$v2Flat] = $numV3Global + $v3TrackIdx
            $v3TrackIdx++
        }
    }

    return $map
}

# ============================================================================
# Build flat paramIndex mapping from OLDER v2 to v3 for Pattern XP data conversion
# Older v2 has different param layout: fewer params, no dividers in tracks
# Returns a hashtable: v2FlatIndex -> v3FlatIndex (or -1 for dividers to skip)
# ============================================================================
function Build-SGridOlderParamIndexMap {
    param([int]$NumChannels)

    $map = @{}
    $droppedColumns = @()

    # --- Older v2 global params (21 params) ---
    # 0: FirstWave
    # 1: Divider
    # 2..(1+N): Triggers 0..(N-1)
    # (2+N): Divider
    # (3+N): TrigType
    # (4+N): Solo
    # (5+N): Vol
    # (6+N): Vel
    # (7+N): HumanVel
    # (8+N): Pan
    # (9+N): HumanPan
    # (10+N): Tune
    # (11+N): HumanTune
    # (12+N): ShufSize  [NOTE: NO LenUnit before this]
    # (13+N): ShufStep
    # (14+N): ShufRnd
    # (15+N): ShufReset
    # (16+N): Inertia
    # Total: N + 17 params (but 21 incl dividers, N=4)

    # v3 global layout (same as before):
    # 0..(N-1): Triggers
    # N: FirstWave
    # (N+1): TrigType
    # (N+2): Solo
    # ...through (N+9): HumanTune
    # (N+10): LenUnit  [v3 has LenUnit here, older doesn't]
    # (N+11): ShufSize
    # ...through (N+14): ShufReset
    # (N+15..N+16): Inertia (word = 1 param but 2 flat indices? No, flat index is per PARAM not per byte)

    # Actually, flat indices in Pattern XP are per PARAMETER not per byte.
    # So Inertia is 1 flat index even though it's 2 bytes.

    # v2 older global flat indices -> v3 global flat indices
    $map[0] = $NumChannels  # FirstWave -> v3[N]
    $map[1] = -1  # Divider -> skip
    for ($i = 0; $i -lt $NumChannels; $i++) {
        $map[2 + $i] = $i  # Triggers -> v3[0..N-1]
    }
    $map[2 + $NumChannels] = -1  # Divider -> skip
    # TrigType..HumanTune (9 params) -> v3 (N+1)..(N+9)
    for ($i = 0; $i -lt 9; $i++) {
        $map[3 + $NumChannels + $i] = $NumChannels + 1 + $i
    }
    # Older has NO LenUnit, so v3[N+10] (LenUnit) has no source
    # ShufSize..ShufReset (4 params) -> v3 (N+11)..(N+14)
    for ($i = 0; $i -lt 4; $i++) {
        $map[3 + $NumChannels + 9 + $i] = $NumChannels + 11 + $i
    }
    # Inertia -> v3 (N+15)
    $map[3 + $NumChannels + 13] = $NumChannels + 15

    $numV2OlderGlobal = $NumChannels + 17  # with 2 dividers: N + 2 + 15 = N+17, but count: 1+1+N+1+9+4+1 = N+17

    # Actually let me recount v2 older globals:
    # FW(1) + Div(1) + Triggers(N) + Div(1) + TrigType(1) + Solo(1) + Vol(1) + Vel(1) +
    # HumanVel(1) + Pan(1) + HumanPan(1) + Tune(1) + HumanTune(1) +
    # ShufSize(1) + ShufStep(1) + ShufRnd(1) + ShufReset(1) + Inertia(1 param)
    # = 1+1+N+1 + 9 + 4 + 1 = N + 17 params. But XML says 21 globals for S4.
    # N=4: 4+17=21. Correct.

    $numV3Global = $NumChannels + 16  # N triggers + FirstWave + 15 common (TrigType..Inertia incl LenUnit)

    # --- Older v2 track params (15 params, NO dividers, NO trigger) ---
    # 0: WaveNo    -> v3: 1
    # 1: Cmd1      -> v3: 2
    # 2: Arg1      -> v3: 3
    # 3: Cmd2      -> v3: 4
    # 4: Arg2      -> v3: 5
    # 5: Subdiv    -> v3: 6 (LenUnit)
    # 6: Mute      -> v3: 7
    # 7: Vol       -> v3: 15
    # 8: Vel       -> v3: 16
    # 9: HumanVel  -> v3: 17
    # 10: Pan      -> v3: 20
    # 11: HumanPan -> v3: 21
    # 12: Tune     -> v3: 24
    # 13: HumanTune-> v3: 25
    # 14: Group    -> v3: 8
    #
    # v3 track (26 params):
    # 0: Trigger, 1: WaveNo, 2: Cmd1, 3: Arg1, 4: Cmd2, 5: Arg2,
    # 6: LenUnit, 7: Mute, 8: Group,
    # 9: Start, 10: LoopStart, 11: End, 12: LoopMode, 13: NoteCut, 14: LoopFit,
    # 15: Vol, 16: Vel, 17: HumanVel, 18: VolEnv, 19: VolEnvLen,
    # 20: Pan, 21: HumanPan, 22: PanEnv, 23: PanEnvLen,
    # 24: Tune, 25: HumanTune

    $v2OlderTrackToV3 = @{
        0 = 1;   # WaveNo
        1 = 2;   # Cmd1
        2 = 3;   # Arg1
        3 = 4;   # Cmd2
        4 = 5;   # Arg2
        5 = 6;   # Subdiv -> LenUnit
        6 = 7;   # Mute
        7 = 15;  # Vol
        8 = 16;  # Vel
        9 = 17;  # HumanVel
        10 = 20; # Pan
        11 = 21; # HumanPan
        12 = 24; # Tune
        13 = 25; # HumanTune
        14 = 8;  # Group
    }

    for ($v2t = 0; $v2t -lt 15; $v2t++) {
        $v2Flat = $numV2OlderGlobal + $v2t
        $map[$v2Flat] = $numV3Global + $v2OlderTrackToV3[$v2t]
    }

    return $map
}

# ============================================================================
# Build flat paramIndex mapping from v1 SWITCH to v3 for Pattern XP conversion
# v1 switch global (no dividers, no TrigType):
#   0:FirstWave, 1..N:Triggers, N+1:Solo, N+2:GlobalVol, N+3:GlobalPan,
#   N+4:GlobalTune, N+5:ShufSize, N+6:ShufStep, N+7:ShufRnd, N+8:ShufReset,
#   N+9:Inertia = N+10 params
#
# v3 global: 0..N-1:Triggers, N:FirstWave, N+1:TrigType, N+2:Solo, N+3:Vol,
#   N+4:Vel, N+5:HumanVel, N+6:Pan, N+7:HumanPan, N+8:Tune, N+9:HumanTune,
#   N+10:LenUnit, N+11:ShufSize..N+14:ShufReset, N+15:Inertia = N+16 params
# ============================================================================
function Build-SGridV1SwitchParamIndexMap {
    param([int]$NumChannels)

    $map = @{}

    # Global mapping
    $map[0] = $NumChannels  # FirstWave -> v3[N]
    for ($i = 0; $i -lt $NumChannels; $i++) {
        $map[1 + $i] = $i  # Triggers -> v3[0..N-1]
    }
    # Solo -> v3[N+2] (skip TrigType at N+1)
    $map[$NumChannels + 1] = $NumChannels + 2
    # GlobalVol -> v3[N+3] (Vol)
    $map[$NumChannels + 2] = $NumChannels + 3
    # GlobalPan -> v3[N+6] (Pan) - skip Vel(N+4), HumanVel(N+5)
    $map[$NumChannels + 3] = $NumChannels + 6
    # GlobalTune -> v3[N+8] (Tune) - skip HumanPan(N+7)
    $map[$NumChannels + 4] = $NumChannels + 8
    # ShufSize..ShufReset -> v3[N+11..N+14] (skip HumanTune(N+9), LenUnit(N+10))
    for ($i = 0; $i -lt 4; $i++) {
        $map[$NumChannels + 5 + $i] = $NumChannels + 11 + $i
    }
    # Inertia -> v3[N+15]
    $map[$NumChannels + 9] = $NumChannels + 15

    $numV1Global = $NumChannels + 10
    $numV3Global = $NumChannels + 16

    # Track mapping (same for v1 switch and v1 byte)
    # v1: 0:WaveNo, 1:Command, 2:Argument, 3:Subdiv, 4:Mute,
    #     5:Volume, 6:Pan, 7:Tune, 8:AuxGroup
    $v1TrackToV3 = @{
        0 = 1;   # WaveNo
        1 = 2;   # Command -> Cmd1
        2 = 3;   # Argument -> Arg1
        3 = 6;   # Subdiv -> LenUnit
        4 = 7;   # Mute
        5 = 15;  # Volume -> Vol
        6 = 20;  # Pan
        7 = 24;  # Tune
        8 = 8;   # AuxGroup -> Group
    }

    for ($v1t = 0; $v1t -lt 9; $v1t++) {
        $v1Flat = $numV1Global + $v1t
        $map[$v1Flat] = $numV3Global + $v1TrackToV3[$v1t]
    }

    return $map
}

# ============================================================================
# Build flat paramIndex mapping from v1 BYTE to v3 for Pattern XP conversion
# v1 byte global (has 2 dividers, has TrigType):
#   0:FirstWave, 1:Div, 2..1+N:Triggers, 2+N:Div, 3+N:TrigType, 4+N:Solo,
#   5+N:GlobalVol, 6+N:GlobalPan, 7+N:GlobalTune,
#   8+N:ShufSize..11+N:ShufReset, 12+N:Inertia = N+13 params
# ============================================================================
function Build-SGridV1ByteParamIndexMap {
    param([int]$NumChannels)

    $map = @{}

    # Global mapping
    $map[0] = $NumChannels  # FirstWave -> v3[N]
    $map[1] = -1  # Divider -> skip
    for ($i = 0; $i -lt $NumChannels; $i++) {
        $map[2 + $i] = $i  # Triggers -> v3[0..N-1]
    }
    $map[2 + $NumChannels] = -1  # Divider -> skip
    # TrigType -> v3[N+1]
    $map[3 + $NumChannels] = $NumChannels + 1
    # Solo -> v3[N+2]
    $map[4 + $NumChannels] = $NumChannels + 2
    # GlobalVol -> v3[N+3]
    $map[5 + $NumChannels] = $NumChannels + 3
    # GlobalPan -> v3[N+6]
    $map[6 + $NumChannels] = $NumChannels + 6
    # GlobalTune -> v3[N+8]
    $map[7 + $NumChannels] = $NumChannels + 8
    # ShufSize..ShufReset -> v3[N+11..N+14]
    for ($i = 0; $i -lt 4; $i++) {
        $map[8 + $NumChannels + $i] = $NumChannels + 11 + $i
    }
    # Inertia -> v3[N+15]
    $map[12 + $NumChannels] = $NumChannels + 15

    $numV1Global = $NumChannels + 13
    $numV3Global = $NumChannels + 16

    # Track mapping (same as v1 switch)
    $v1TrackToV3 = @{
        0 = 1; 1 = 2; 2 = 3; 3 = 6; 4 = 7; 5 = 15; 6 = 20; 7 = 24; 8 = 8
    }
    for ($v1t = 0; $v1t -lt 9; $v1t++) {
        $v1Flat = $numV1Global + $v1t
        $map[$v1Flat] = $numV3Global + $v1TrackToV3[$v1t]
    }

    return $map
}

# ============================================================================
# Build flat paramIndex mapping from v2 MID (050604) to v3 for Pattern XP
# Globals: same layout as older v2 (N+17 params with 2 dividers)
# Tracks (20 params, no dividers):
#   0:Trigger, 1:WaveNo, 2:Cmd1, 3:Arg1, 4:Cmd2, 5:Arg2, 6:Subdiv, 7:Mute,
#   8:Offset, 9:NoteCut, 10:Vol, 11:Vel, 12:HumanVel, 13:Pan, 14:HumanPan,
#   15:Tune, 16:HumanTune, 17:LoopFit, 18:LpFitMode, 19:Group
# ============================================================================
function Build-SGridMidParamIndexMap {
    param([int]$NumChannels)

    $map = @{}

    # Globals: same as older v2 (FW, Div, Trigs, Div, TrigType..Inertia minus LenUnit)
    $map[0] = $NumChannels  # FirstWave
    $map[1] = -1  # Divider
    for ($i = 0; $i -lt $NumChannels; $i++) { $map[2 + $i] = $i }
    $map[2 + $NumChannels] = -1  # Divider
    for ($i = 0; $i -lt 9; $i++) {
        $map[3 + $NumChannels + $i] = $NumChannels + 1 + $i  # TrigType..HumanTune
    }
    for ($i = 0; $i -lt 4; $i++) {
        $map[3 + $NumChannels + 9 + $i] = $NumChannels + 11 + $i  # ShufSize..ShufReset (skip LenUnit)
    }
    $map[3 + $NumChannels + 13] = $NumChannels + 15  # Inertia

    $numV2Global = $NumChannels + 17
    $numV3Global = $NumChannels + 16

    # Track mapping
    # v2 mid -> v3 track param indices
    $midTrackToV3 = @{
        0 = 0;   # Trigger
        1 = 1;   # WaveNo
        2 = 2;   # Cmd1
        3 = 3;   # Arg1
        4 = 4;   # Cmd2
        5 = 5;   # Arg2
        6 = 6;   # Subdiv -> LenUnit
        7 = 7;   # Mute
        8 = 9;   # Offset -> Start
        9 = 13;  # NoteCut
        10 = 15; # Vol
        11 = 16; # Vel
        12 = 17; # HumanVel
        13 = 20; # Pan
        14 = 21; # HumanPan
        15 = 24; # Tune
        16 = 25; # HumanTune
        17 = 14; # LoopFit
        18 = -1; # LpFitMode -> DROPPED
        19 = 8;  # Group
    }

    for ($v2t = 0; $v2t -lt 20; $v2t++) {
        $v2Flat = $numV2Global + $v2t
        $v3Target = $midTrackToV3[$v2t]
        if ($v3Target -eq -1) {
            $map[$v2Flat] = -1  # dropped
        } else {
            $map[$v2Flat] = $numV3Global + $v3Target
        }
    }

    return $map
}

# ============================================================================
# Build flat paramIndex mapping from v2 LATE (050716) to v3 for Pattern XP
# Globals: same layout as older v2
# Tracks (26 params, no dividers):
#   0:Trigger, 1:WaveNo, 2:Cmd1, 3:Arg1, 4:Cmd2, 5:Arg2, 6:Subdiv, 7:Mute,
#   8:Offset, 9:NoteCut, 10:Vol, 11:Vel, 12:HumanVel, 13:VolEnv, 14:VolEnvLen,
#   15:Pan, 16:HumanPan, 17:PanEnv, 18:PanEnvLen, 19:Tune, 20:HumanTune,
#   21:TuneEnv, 22:TuneEnvLen, 23:LoopFit, 24:LpFitMode, 25:Group
# ============================================================================
function Build-SGridLateParamIndexMap {
    param([int]$NumChannels)

    $map = @{}

    # Globals: same as older v2
    $map[0] = $NumChannels
    $map[1] = -1
    for ($i = 0; $i -lt $NumChannels; $i++) { $map[2 + $i] = $i }
    $map[2 + $NumChannels] = -1
    for ($i = 0; $i -lt 9; $i++) {
        $map[3 + $NumChannels + $i] = $NumChannels + 1 + $i
    }
    for ($i = 0; $i -lt 4; $i++) {
        $map[3 + $NumChannels + 9 + $i] = $NumChannels + 11 + $i
    }
    $map[3 + $NumChannels + 13] = $NumChannels + 15

    $numV2Global = $NumChannels + 17
    $numV3Global = $NumChannels + 16

    # Track mapping
    $lateTrackToV3 = @{
        0 = 0;   # Trigger
        1 = 1;   # WaveNo
        2 = 2;   # Cmd1
        3 = 3;   # Arg1
        4 = 4;   # Cmd2
        5 = 5;   # Arg2
        6 = 6;   # Subdiv -> LenUnit
        7 = 7;   # Mute
        8 = 9;   # Offset -> Start
        9 = 13;  # NoteCut
        10 = 15; # Vol
        11 = 16; # Vel
        12 = 17; # HumanVel
        13 = 18; # VolEnv
        14 = 19; # VolEnvLen
        15 = 20; # Pan
        16 = 21; # HumanPan
        17 = 22; # PanEnv
        18 = 23; # PanEnvLen
        19 = 24; # Tune
        20 = 25; # HumanTune
        21 = -1; # TuneEnv -> DROPPED
        22 = -1; # TuneEnvLen -> DROPPED
        23 = 14; # LoopFit
        24 = -1; # LpFitMode -> DROPPED
        25 = 8;  # Group
    }

    for ($v2t = 0; $v2t -lt 26; $v2t++) {
        $v2Flat = $numV2Global + $v2t
        $v3Target = $lateTrackToV3[$v2t]
        if ($v3Target -eq -1) {
            $map[$v2Flat] = -1
        } else {
            $map[$v2Flat] = $numV3Global + $v3Target
        }
    }

    return $map
}

# ============================================================================
# Remap PAT2 column indices: v1 SWITCH param indices -> v3 param indices
# v1 switch has NO dividers in globals and NO TrigType
# ============================================================================
function Convert-SGridV1SwitchPat2ColumnIndex {
    param(
        [int]$Group,
        [int]$IndexInGroup,
        [int]$NumChannels
    )

    if ($Group -eq 1) {
        # v1 switch global (no dividers):
        # 0:FirstWave, 1..N:Triggers, N+1:Solo, N+2:GlobalVol, N+3:GlobalPan,
        # N+4:GlobalTune, N+5:ShufSize, N+6:ShufStep, N+7:ShufRnd,
        # N+8:ShufReset, N+9:Inertia

        if ($IndexInGroup -eq 0) { return @{ skip = $false; newIndex = $NumChannels } }  # FirstWave -> N
        if ($IndexInGroup -ge 1 -and $IndexInGroup -le $NumChannels) {
            return @{ skip = $false; newIndex = ($IndexInGroup - 1) }  # Triggers -> 0..N-1
        }
        # After triggers: Solo, GlobalVol, GlobalPan, GlobalTune, ShufSize..Inertia
        # v3 order after FirstWave: TrigType(N+1), Solo(N+2), Vol(N+3), Vel(N+4),
        #   HumanVel(N+5), Pan(N+6), HumanPan(N+7), Tune(N+8), HumanTune(N+9),
        #   LenUnit(N+10), ShufSize(N+11)..ShufReset(N+14), Inertia(N+15)
        $posAfterTrigs = $IndexInGroup - $NumChannels - 1  # 0=Solo, 1=GlobalVol, 2=GlobalPan, 3=GlobalTune
        switch ($posAfterTrigs) {
            0 { return @{ skip = $false; newIndex = $NumChannels + 2 } }   # Solo
            1 { return @{ skip = $false; newIndex = $NumChannels + 3 } }   # GlobalVol -> Vol
            2 { return @{ skip = $false; newIndex = $NumChannels + 6 } }   # GlobalPan -> Pan
            3 { return @{ skip = $false; newIndex = $NumChannels + 8 } }   # GlobalTune -> Tune
            { $_ -ge 4 -and $_ -le 7 } { return @{ skip = $false; newIndex = $NumChannels + 7 + $posAfterTrigs } }  # ShufSize..ShufReset -> N+11..N+14
            8 { return @{ skip = $false; newIndex = $NumChannels + 15 } }  # Inertia
        }
        return @{ skip = $false; newIndex = $IndexInGroup }
    }
    elseif ($Group -eq 2) {
        # v1 track (9 params, no dividers):
        # 0:WaveNo, 1:Command, 2:Argument, 3:Subdiv, 4:Mute,
        # 5:Volume, 6:Pan, 7:Tune, 8:AuxGroup
        $trackMap = @{
            0 = 1; 1 = 2; 2 = 3; 3 = 6; 4 = 7;
            5 = 15; 6 = 20; 7 = 24; 8 = 8
        }
        if ($trackMap.ContainsKey($IndexInGroup)) {
            return @{ skip = $false; newIndex = $trackMap[$IndexInGroup] }
        }
        return @{ skip = $false; newIndex = $IndexInGroup }
    }

    return @{ skip = $false; newIndex = $IndexInGroup }
}

# ============================================================================
# Remap PAT2 column indices: v1 BYTE param indices -> v3 param indices
# v1 byte HAS 2 dividers in globals and HAS TrigType
# ============================================================================
function Convert-SGridV1BytePat2ColumnIndex {
    param(
        [int]$Group,
        [int]$IndexInGroup,
        [int]$NumChannels
    )

    if ($Group -eq 1) {
        # v1 byte global (with dividers):
        # 0:FirstWave, 1:Div, 2..1+N:Triggers, 2+N:Div, 3+N:TrigType,
        # 4+N:Solo, 5+N:GlobalVol, 6+N:GlobalPan, 7+N:GlobalTune,
        # 8+N:ShufSize..11+N:ShufReset, 12+N:Inertia

        if ($IndexInGroup -eq 0) { return @{ skip = $false; newIndex = $NumChannels } }  # FirstWave
        if ($IndexInGroup -eq 1) { return @{ skip = $true; newIndex = -1 } }  # Divider
        if ($IndexInGroup -ge 2 -and $IndexInGroup -le (1 + $NumChannels)) {
            return @{ skip = $false; newIndex = ($IndexInGroup - 2) }  # Triggers
        }
        if ($IndexInGroup -eq (2 + $NumChannels)) { return @{ skip = $true; newIndex = -1 } }  # Divider

        $posAfterDiv2 = $IndexInGroup - (3 + $NumChannels)  # 0=TrigType, 1=Solo, 2=GlobalVol...
        switch ($posAfterDiv2) {
            0 { return @{ skip = $false; newIndex = $NumChannels + 1 } }   # TrigType
            1 { return @{ skip = $false; newIndex = $NumChannels + 2 } }   # Solo
            2 { return @{ skip = $false; newIndex = $NumChannels + 3 } }   # GlobalVol -> Vol
            3 { return @{ skip = $false; newIndex = $NumChannels + 6 } }   # GlobalPan -> Pan
            4 { return @{ skip = $false; newIndex = $NumChannels + 8 } }   # GlobalTune -> Tune
            { $_ -ge 5 -and $_ -le 8 } { return @{ skip = $false; newIndex = $NumChannels + 6 + $posAfterDiv2 } }  # ShufSize..ShufReset -> N+11..N+14
            9 { return @{ skip = $false; newIndex = $NumChannels + 15 } }  # Inertia
        }
        return @{ skip = $false; newIndex = $IndexInGroup }
    }
    elseif ($Group -eq 2) {
        # Same track mapping as v1 switch
        $trackMap = @{
            0 = 1; 1 = 2; 2 = 3; 3 = 6; 4 = 7;
            5 = 15; 6 = 20; 7 = 24; 8 = 8
        }
        if ($trackMap.ContainsKey($IndexInGroup)) {
            return @{ skip = $false; newIndex = $trackMap[$IndexInGroup] }
        }
        return @{ skip = $false; newIndex = $IndexInGroup }
    }

    return @{ skip = $false; newIndex = $IndexInGroup }
}

# ============================================================================
# Remap PAT2 column indices: v2 MID (050604) param indices -> v3 param indices
# Globals: same as older v2 (has dividers, no LenUnit)
# Tracks: 20 params, no dividers
# ============================================================================
function Convert-SGridMidPat2ColumnIndex {
    param(
        [int]$Group,
        [int]$IndexInGroup,
        [int]$NumChannels
    )

    if ($Group -eq 1) {
        # Same as older v2 globals (Convert-SGridOlderPat2ColumnIndex)
        return Convert-SGridOlderPat2ColumnIndex -Group $Group -IndexInGroup $IndexInGroup -NumChannels $NumChannels
    }
    elseif ($Group -eq 2) {
        # v2 mid track (20 params):
        # 0:Trigger, 1:WaveNo, 2:Cmd1, 3:Arg1, 4:Cmd2, 5:Arg2, 6:Subdiv, 7:Mute,
        # 8:Offset, 9:NoteCut, 10:Vol, 11:Vel, 12:HumanVel, 13:Pan, 14:HumanPan,
        # 15:Tune, 16:HumanTune, 17:LoopFit, 18:LpFitMode, 19:Group
        $trackMap = @{
            0 = 0;   # Trigger
            1 = 1;   # WaveNo
            2 = 2;   # Cmd1
            3 = 3;   # Arg1
            4 = 4;   # Cmd2
            5 = 5;   # Arg2
            6 = 6;   # Subdiv -> LenUnit
            7 = 7;   # Mute
            8 = 9;   # Offset -> Start
            9 = 13;  # NoteCut
            10 = 15; # Vol
            11 = 16; # Vel
            12 = 17; # HumanVel
            13 = 20; # Pan
            14 = 21; # HumanPan
            15 = 24; # Tune
            16 = 25; # HumanTune
            17 = 14; # LoopFit
            19 = 8;  # Group
        }
        if ($IndexInGroup -eq 18) { return @{ skip = $true; newIndex = -1 } }  # LpFitMode -> DROPPED
        if ($trackMap.ContainsKey($IndexInGroup)) {
            return @{ skip = $false; newIndex = $trackMap[$IndexInGroup] }
        }
        return @{ skip = $false; newIndex = $IndexInGroup }
    }

    return @{ skip = $false; newIndex = $IndexInGroup }
}

# ============================================================================
# Remap PAT2 column indices: v2 LATE (050716) param indices -> v3 param indices
# Globals: same as older v2
# Tracks: 26 params, no dividers
# ============================================================================
function Convert-SGridLatePat2ColumnIndex {
    param(
        [int]$Group,
        [int]$IndexInGroup,
        [int]$NumChannels
    )

    if ($Group -eq 1) {
        return Convert-SGridOlderPat2ColumnIndex -Group $Group -IndexInGroup $IndexInGroup -NumChannels $NumChannels
    }
    elseif ($Group -eq 2) {
        # v2 late track (26 params):
        # 0:Trigger, 1:WaveNo, 2:Cmd1, 3:Arg1, 4:Cmd2, 5:Arg2, 6:Subdiv, 7:Mute,
        # 8:Offset, 9:NoteCut, 10:Vol, 11:Vel, 12:HumanVel, 13:VolEnv, 14:VolEnvLen,
        # 15:Pan, 16:HumanPan, 17:PanEnv, 18:PanEnvLen, 19:Tune, 20:HumanTune,
        # 21:TuneEnv, 22:TuneEnvLen, 23:LoopFit, 24:LpFitMode, 25:Group
        $trackMap = @{
            0 = 0; 1 = 1; 2 = 2; 3 = 3; 4 = 4; 5 = 5; 6 = 6; 7 = 7;
            8 = 9; 9 = 13; 10 = 15; 11 = 16; 12 = 17; 13 = 18; 14 = 19;
            15 = 20; 16 = 21; 17 = 22; 18 = 23; 19 = 24; 20 = 25;
            23 = 14; 25 = 8
        }
        # Dropped params: 21=TuneEnv, 22=TuneEnvLen, 24=LpFitMode
        if ($IndexInGroup -in @(21, 22, 24)) { return @{ skip = $true; newIndex = -1 } }
        if ($trackMap.ContainsKey($IndexInGroup)) {
            return @{ skip = $false; newIndex = $trackMap[$IndexInGroup] }
        }
        return @{ skip = $false; newIndex = $IndexInGroup }
    }

    return @{ skip = $false; newIndex = $IndexInGroup }
}

# ============================================================================
# Convert Pattern XP data blob: remap paramIndex values from v2 to v3
# and remove divider columns entirely.
# The blob format (version 3):
#   version(1b), numPatterns(int32)
#   per pattern: name(asciiz), rowsPerBeat(int32), numColumns(int32)
#   per column: machineName(asciiz), paramIndex(int32), paramTrack(int32),
#               graphical(1b), numEvents(int32)
#   per event: row(int32), value(int32)
#
# Since divider columns must be removed (changing the size), we rebuild the
# blob into a new MemoryStream rather than editing in-place.
# ============================================================================
function Convert-PatternXPData {
    param(
        [byte[]]$Data,
        [string]$SGridMachineName,  # name of the SGrid machine this editor targets
        [int]$NumChannels,
        [string]$Revision = "061110"
    )

    if ($Data.Length -eq 0) { return $Data }

    # Build index map based on revision
    switch ($Revision) {
        "061110"   { $indexMap = Build-SGridParamIndexMap -NumChannels $NumChannels }          # Build-SGridParamIndexMap in SampleGridUpgrade.ps1
        "older"    { $indexMap = Build-SGridOlderParamIndexMap -NumChannels $NumChannels }     # Build-SGridOlderParamIndexMap in SampleGridUpgrade.ps1
        "mid"      { $indexMap = Build-SGridMidParamIndexMap -NumChannels $NumChannels }       # Build-SGridMidParamIndexMap in SampleGridUpgrade.ps1
        "late"     { $indexMap = Build-SGridLateParamIndexMap -NumChannels $NumChannels }      # Build-SGridLateParamIndexMap in SampleGridUpgrade.ps1
        "v1switch" { $indexMap = Build-SGridV1SwitchParamIndexMap -NumChannels $NumChannels }  # Build-SGridV1SwitchParamIndexMap in SampleGridUpgrade.ps1
        "v1byte"   { $indexMap = Build-SGridV1ByteParamIndexMap -NumChannels $NumChannels }    # Build-SGridV1ByteParamIndexMap in SampleGridUpgrade.ps1
        default    { $indexMap = Build-SGridOlderParamIndexMap -NumChannels $NumChannels }
    }

    $pos = 0

    # --- Read helpers ---
    function Read-Asciiz {
        param([byte[]]$buf, [ref]$p)
        $start = $p.Value
        while ($p.Value -lt $buf.Length -and $buf[$p.Value] -ne 0) { $p.Value++ }
        $str = [System.Text.Encoding]::ASCII.GetString($buf, $start, $p.Value - $start)
        $p.Value++  # skip null
        return $str
    }

    function Read-I32 {
        param([byte[]]$buf, [ref]$p)
        $val = [BitConverter]::ToInt32($buf, $p.Value)
        $p.Value += 4
        return $val
    }

    # --- Write helpers ---
    function Out-Asciiz {
        param([System.IO.MemoryStream]$ms, [string]$str)
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($str)
        $ms.Write($bytes, 0, $bytes.Length)
        $ms.WriteByte(0)
    }

    function Out-I32 {
        param([System.IO.MemoryStream]$ms, [int]$val)
        $bytes = [BitConverter]::GetBytes([int]$val)
        $ms.Write($bytes, 0, 4)
    }

    function Out-Byte {
        param([System.IO.MemoryStream]$ms, [byte]$val)
        $ms.WriteByte($val)
    }

    $version = $Data[$pos]; $pos++
    if ($version -lt 1 -or $version -gt 3) {
        Write-Log "    WARNING: Unknown Pattern XP data version $version, skipping conversion"
        return $Data
    }

    $numPatterns = Read-I32 $Data ([ref]$pos)

    # Rebuild the blob into a new stream
    $out = New-Object System.IO.MemoryStream
    Out-Byte $out $version
    Out-I32 $out $numPatterns

    $totalRemoved = 0

    for ($p = 0; $p -lt $numPatterns; $p++) {
        # Pattern name
        $patName = Read-Asciiz $Data ([ref]$pos)
        Out-Asciiz $out $patName

        # rowsPerBeat (version > 1)
        if ($version -gt 1) {
            $rowsPerBeat = Read-I32 $Data ([ref]$pos)
            Out-I32 $out $rowsPerBeat
        }

        # Read all columns first, then write only non-divider ones
        $numColumns = Read-I32 $Data ([ref]$pos)

        # Parse all columns into a list
        $columnList = @()
        for ($c = 0; $c -lt $numColumns; $c++) {
            $col = @{}
            $col.machineName = Read-Asciiz $Data ([ref]$pos)
            $col.paramIndex = Read-I32 $Data ([ref]$pos)
            $col.paramTrack = Read-I32 $Data ([ref]$pos)

            if ($version -ge 3) {
                $col.graphical = $Data[$pos]; $pos++
            }

            $col.numEvents = Read-I32 $Data ([ref]$pos)
            $col.events = @()
            for ($e = 0; $e -lt $col.numEvents; $e++) {
                $row = Read-I32 $Data ([ref]$pos)
                $value = Read-I32 $Data ([ref]$pos)
                $col.events += ,@($row, $value)
            }

            $columnList += ,$col
        }

        # Remap SGrid columns and rebuild in v3 order
        # First, separate SGrid columns from non-SGrid columns
        $sgridCols = @()
        $nonSgridCols = @()
        foreach ($col in $columnList) {
            if ($col.machineName -eq $SGridMachineName) {
                if ($indexMap.ContainsKey($col.paramIndex)) {
                    $newIndex = $indexMap[$col.paramIndex]
                    if ($newIndex -eq -1) {
                        # Divider column - remove it
                        $totalRemoved++
                        continue
                    }
                    $col.paramIndex = $newIndex
                } else {
                    Write-Log "    WARNING: PatternXP paramIndex $($col.paramIndex) has no mapping"
                }
                $sgridCols += ,$col
            } else {
                $nonSgridCols += ,$col
            }
        }

        # Build a lookup of remapped SGrid columns by (paramIndex, paramTrack)
        $sgridLookup = @{}
        foreach ($col in $sgridCols) {
            $key = "$($col.paramIndex),$($col.paramTrack)"
            $sgridLookup[$key] = $col
        }

        # Determine numTracks from existing columns
        $maxTrack = -1
        foreach ($col in $sgridCols) {
            if ($col.paramTrack -gt $maxTrack) { $maxTrack = $col.paramTrack }
        }
        $peNumTracks = $maxTrack + 1

        # Rebuild SGrid columns in v3 order with all params present
        $numV3Global = $NumChannels + 16
        $numV3Track = 26
        $rebuiltSgridCols = @()

        # Global params (flat index 0..numV3Global-1, track=0)
        for ($gi = 0; $gi -lt $numV3Global; $gi++) {
            $key = "$gi,0"
            if ($sgridLookup.ContainsKey($key)) {
                $rebuiltSgridCols += ,$sgridLookup[$key]
            } else {
                # New v3 param - add empty column
                $newCol = @{
                    machineName = $SGridMachineName
                    paramIndex = $gi
                    paramTrack = 0
                    numEvents = 0
                    events = @()
                }
                if ($version -ge 3) { $newCol.graphical = 0 }
                $rebuiltSgridCols += ,$newCol
            }
        }

        # Track params: paramIndex = numV3Global + offsetInTrack (same for all tracks)
        for ($t = 0; $t -lt $peNumTracks; $t++) {
            for ($ti = 0; $ti -lt $numV3Track; $ti++) {
                $flatIdx = $numV3Global + $ti
                $key = "$flatIdx,$t"
                if ($sgridLookup.ContainsKey($key)) {
                    $rebuiltSgridCols += ,$sgridLookup[$key]
                } else {
                    $newCol = @{
                        machineName = $SGridMachineName
                        paramIndex = $flatIdx
                        paramTrack = $t
                        numEvents = 0
                        events = @()
                    }
                    if ($version -ge 3) { $newCol.graphical = 0 }
                    $rebuiltSgridCols += ,$newCol
                }
            }
        }

        # Combine: rebuilt SGrid columns first, then non-SGrid columns
        $outputColumns = $rebuiltSgridCols + $nonSgridCols

        # Write column count and columns
        Out-I32 $out $outputColumns.Count

        foreach ($col in $outputColumns) {
            Out-Asciiz $out $col.machineName
            Out-I32 $out $col.paramIndex
            Out-I32 $out $col.paramTrack

            if ($version -ge 3) {
                Out-Byte $out ([byte]$col.graphical)
            }

            Out-I32 $out $col.events.Count

            foreach ($ev in $col.events) {
                Out-I32 $out $ev[0]  # row
                Out-I32 $out $ev[1]  # value
            }
        }
    }

    Write-Log "    Removed $totalRemoved divider columns from Pattern XP data"
    $result = $out.ToArray()
    $out.Close()
    return $result
}

# ============================================================================
# Remap PAT2 column indices: v2 061110 param indices -> v3 param indices
# PAT2 columns reference (group, indexInGroup) where group 1=global, group 2=track
# For global params: remap indices to account for removed dividers and reordering
# For track params: remap indices to account for removed dividers
# ============================================================================
function Convert-SGridPat2ColumnIndex {
    param(
        [int]$Group,         # 1=global, 2=track
        [int]$IndexInGroup,  # 0-based param index within group
        [int]$NumChannels    # 4, 8, 16, or 32
    )

    if ($Group -eq 1) {
        # Global param index remapping
        # v2 order (0-based): 0=FirstWave, 1=Divider, 2..1+N=Triggers, 2+N=Divider, 3+N..end=rest
        # v3 order (0-based): 0..N-1=Triggers, N=FirstWave, N+1..end=rest

        if ($IndexInGroup -eq 0) {
            # FirstWave -> moves to position N
            return @{ skip = $false; newIndex = $NumChannels }
        }
        if ($IndexInGroup -eq 1) {
            # Divider -> skip
            return @{ skip = $true; newIndex = -1 }
        }
        if ($IndexInGroup -ge 2 -and $IndexInGroup -le (1 + $NumChannels)) {
            # Trigger i -> moves to position (i - 2)
            return @{ skip = $false; newIndex = ($IndexInGroup - 2) }
        }
        if ($IndexInGroup -eq (2 + $NumChannels)) {
            # Divider -> skip
            return @{ skip = $true; newIndex = -1 }
        }
        # Rest params: shift down by 2 (removed 2 dividers)
        return @{ skip = $false; newIndex = ($IndexInGroup - 2) }
    }
    elseif ($Group -eq 2) {
        # Track param index remapping: just remove dividers
        # v2 divider indices (0-based): 1, 10, 17, 23, 28
        $dividers = $script:SGridV2TrackDividerIndices
        if ($IndexInGroup -in $dividers) {
            return @{ skip = $true; newIndex = -1 }
        }
        # Count how many dividers come before this index
        $shift = 0
        foreach ($d in $dividers) {
            if ($d -lt $IndexInGroup) { $shift++ }
        }
        return @{ skip = $false; newIndex = ($IndexInGroup - $shift) }
    }

    # Group 0 or other (input connections) - no change
    return @{ skip = $false; newIndex = $IndexInGroup }
}

# ============================================================================
# Remap PAT2 column indices: OLDER v2 param indices -> v3 param indices
# Older v2 has different param layout - no dividers in tracks, fewer params,
# and different ordering.
# ============================================================================
function Convert-SGridOlderPat2ColumnIndex {
    param(
        [int]$Group,         # 1=global, 2=track
        [int]$IndexInGroup,  # 0-based param index within group
        [int]$NumChannels    # 4, 8, 16, or 32
    )

    if ($Group -eq 1) {
        # Older v2 global params (same divider layout as 061110 for globals):
        # 0=FirstWave, 1=Divider, 2..1+N=Triggers, 2+N=Divider, 3+N..end=rest
        # But the rest is slightly different: NO LenUnit
        # v2 older rest: TrigType, Solo, Vol, Vel, HumanVel, Pan, HumanPan, Tune, HumanTune,
        #   ShufSize, ShufStep, ShufRnd, ShufReset, Inertia  [14 params]
        # v3 rest: TrigType, Solo, Vol, Vel, HumanVel, Pan, HumanPan, Tune, HumanTune,
        #   LenUnit, ShufSize, ShufStep, ShufRnd, ShufReset, Inertia  [15 params]

        if ($IndexInGroup -eq 0) {
            return @{ skip = $false; newIndex = $NumChannels }  # FirstWave -> N
        }
        if ($IndexInGroup -eq 1) {
            return @{ skip = $true; newIndex = -1 }  # Divider -> skip
        }
        if ($IndexInGroup -ge 2 -and $IndexInGroup -le (1 + $NumChannels)) {
            return @{ skip = $false; newIndex = ($IndexInGroup - 2) }  # Triggers -> 0..N-1
        }
        if ($IndexInGroup -eq (2 + $NumChannels)) {
            return @{ skip = $true; newIndex = -1 }  # Divider -> skip
        }
        # Rest params after 2nd divider: position in v2 = IndexInGroup - (3 + N)
        $posInRest = $IndexInGroup - (3 + $NumChannels)
        if ($posInRest -lt 9) {
            # TrigType..HumanTune (first 9 of rest) -> v3 (N+1)..(N+9)
            return @{ skip = $false; newIndex = $NumChannels + 1 + $posInRest }
        }
        # After HumanTune, older has ShufSize; v3 has LenUnit then ShufSize
        # So v2 rest[9..12] (ShufSize..ShufReset) -> v3 (N+11)..(N+14) [shifted +1 for LenUnit]
        if ($posInRest -ge 9 -and $posInRest -le 12) {
            return @{ skip = $false; newIndex = $NumChannels + 2 + $posInRest }  # +2 because LenUnit inserted
        }
        # v2 rest[13] = Inertia -> v3 (N+15)
        if ($posInRest -eq 13) {
            return @{ skip = $false; newIndex = $NumChannels + 15 }
        }
        # Shouldn't reach here
        return @{ skip = $false; newIndex = $IndexInGroup }
    }
    elseif ($Group -eq 2) {
        # Older v2 track params (15 params, NO dividers):
        # 0:WaveNo, 1:Cmd1, 2:Arg1, 3:Cmd2, 4:Arg2, 5:Subdiv, 6:Mute,
        # 7:Vol, 8:Vel, 9:HumanVel, 10:Pan, 11:HumanPan, 12:Tune, 13:HumanTune, 14:Group
        #
        # v3 track params (26 params):
        # 0:Trigger, 1:WaveNo, 2:Cmd1, 3:Arg1, 4:Cmd2, 5:Arg2, 6:LenUnit, 7:Mute, 8:Group,
        # 9:Start, 10:LoopStart, 11:End, 12:LoopMode, 13:NoteCut, 14:LoopFit,
        # 15:Vol, 16:Vel, 17:HumanVel, 18:VolEnv, 19:VolEnvLen,
        # 20:Pan, 21:HumanPan, 22:PanEnv, 23:PanEnvLen, 24:Tune, 25:HumanTune

        $trackMap = @{
            0 = 1;   # WaveNo
            1 = 2;   # Cmd1
            2 = 3;   # Arg1
            3 = 4;   # Cmd2
            4 = 5;   # Arg2
            5 = 6;   # Subdiv -> LenUnit
            6 = 7;   # Mute
            7 = 15;  # Vol
            8 = 16;  # Vel
            9 = 17;  # HumanVel
            10 = 20; # Pan
            11 = 21; # HumanPan
            12 = 24; # Tune
            13 = 25; # HumanTune
            14 = 8;  # Group
        }

        if ($trackMap.ContainsKey($IndexInGroup)) {
            return @{ skip = $false; newIndex = $trackMap[$IndexInGroup] }
        }
        return @{ skip = $false; newIndex = $IndexInGroup }
    }

    # Group 0 or other (input connections) - no change
    return @{ skip = $false; newIndex = $IndexInGroup }
}

# ============================================================================
# Main upgrade function: decode -> modify XML -> re-encode
# ============================================================================
function Invoke-UpgradeSampleGrid {
    param(
        [string]$BmxPath,
        [string]$OutPath
    )

    Write-Log "Starting SampleGrid upgrade: $BmxPath -> $OutPath"

    # Step 1: Decode to temp XML
    $tempXml = [System.IO.Path]::GetTempFileName()
    $tempXml = [System.IO.Path]::ChangeExtension($tempXml, ".xml")
    Write-Log "  Decoding to temp XML: $tempXml"
    ConvertFrom-Buzz -BmxPath $BmxPath -XmlPath $tempXml  # decode (ConvertFrom-Buzz)

    # Step 2: Load XML
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $false
    $xml.Load($tempXml)
    $root = $xml.DocumentElement

    $paraEl = $root.SelectSingleNode("PARA")
    $machEl = $root.SelectSingleNode("MACH")
    $pattEl = $root.SelectSingleNode("PATT")
    $pat2El = $root.SelectSingleNode("PAT2")
    $patxEl = $root.SelectSingleNode("PATX")

    if (-not $paraEl) { throw "No PARA section found in decoded XML" }

    # Step 3: Find all upgradeable SampleGrid machines (v1 and v2)
    $sgMachines = @()
    foreach ($mach in @($paraEl.SelectNodes("Machine"))) {
        $machType = $mach.GetAttribute("type")
        if (Test-SGridUpgradeable $machType) {  # Test-SGridUpgradeable in SampleGridUpgrade.ps1
            $numTrackParams = [int]$mach.GetAttribute("numTrackParams")
            $revision = Get-SGridRevision -TypeString $machType -NumTrackParams $numTrackParams  # Get-SGridRevision in SampleGridUpgrade.ps1
            $sgMachines += @{
                Name = $mach.GetAttribute("name")
                Type = $machType
                NumChannels = Get-SGridChannelCount $machType  # Get-SGridChannelCount in SampleGridUpgrade.ps1
                IsBVariant = Test-SGridBVariant $machType  # Test-SGridBVariant in SampleGridUpgrade.ps1
                Revision = $revision
            }
        }
    }

    if ($sgMachines.Count -eq 0) {
        Write-Host "No BTDSys SampleGrid machines found in file. Nothing to upgrade."
        Write-Log "No upgradeable SampleGrid machines found."
        Remove-Item $tempXml -Force -ErrorAction SilentlyContinue
        return
    }

    # Revision label for display
    $revisionLabels = @{
        "061110" = "v2 061110"
        "older" = "v2 older (050422)"
        "mid" = "v2 mid (050604)"
        "late" = "v2 late (050716)"
        "v1switch" = "v1 switch"
        "v1byte" = "v1 byte"
    }

    Write-Host "Found $($sgMachines.Count) BTDSys SampleGrid machine(s) to upgrade:"
    foreach ($m in $sgMachines) {
        $variant = if ($m.IsBVariant) { "B" } else { "S" }
        $revLabel = if ($revisionLabels.ContainsKey($m.Revision)) { $revisionLabels[$m.Revision] } else { $m.Revision }
        Write-Host "  $($m.Name) ($variant$($m.NumChannels), $revLabel)"
    }

    # Show warnings about missing parameters per revision
    $revisionsMissing = @{
        "v1switch" = @{
            Track = "Trigger, Cmd2, Arg2, Start, LoopStart, End, LoopMode, NoteCut, LoopFit, Velocity, HumanVel, VolEnv, VolEnvLen, HumanPan, PanEnv, PanEnvLen, HumanTune"
            Global = "TrigType, Velocity, HumanVel, HumanPan, HumanTune, LenUnit"
        }
        "v1byte" = @{
            Track = "Trigger, Cmd2, Arg2, Start, LoopStart, End, LoopMode, NoteCut, LoopFit, Velocity, HumanVel, VolEnv, VolEnvLen, HumanPan, PanEnv, PanEnvLen, HumanTune"
            Global = "Velocity, HumanVel, HumanPan, HumanTune, LenUnit"
        }
        "older" = @{
            Track = "Trigger, Start, LoopStart, End, LoopMode, NoteCut, LoopFit, VolEnv, VolEnvLen, PanEnv, PanEnvLen"
            Global = "LenUnit"
        }
        "mid" = @{
            Track = "LoopStart, End, LoopMode, VolEnv, VolEnvLen, PanEnv, PanEnvLen"
            Global = "LenUnit"
            Dropped = "LpFitMode"
        }
        "late" = @{
            Track = "LoopStart, End, LoopMode"
            Global = "LenUnit"
            Dropped = "TuneEnv, TuneEnvLen, LpFitMode"
        }
    }

    $uniqueRevisions = $sgMachines | ForEach-Object { $_.Revision } | Sort-Object -Unique
    foreach ($rev in $uniqueRevisions) {
        if ($rev -ne "061110" -and $revisionsMissing.ContainsKey($rev)) {
            $missing = $revisionsMissing[$rev]
            $revLabel = $revisionLabels[$rev]
            Write-Host ""
            Write-Host "NOTE: $revLabel machines have fewer parameters than v3."
            if ($missing.Track) {
                Write-Host "      Missing track params (set to defaults): $($missing.Track)"
            }
            if ($missing.Global) {
                Write-Host "      Missing global params (set to defaults): $($missing.Global)"
            }
            if ($missing.Dropped) {
                Write-Host "      Dropped params (not in v3): $($missing.Dropped)"
            }
        }
    }
    Write-Host ""

    # Build lookup of machine names for quick access
    $v2NameSet = @{}
    foreach ($m in $sgMachines) { $v2NameSet[$m.Name] = $m }

    # ========================================================================
    # Step 4: Upgrade PARA section - replace params
    # ========================================================================
    Write-Log "  Upgrading PARA section..."
    foreach ($mach in @($paraEl.SelectNodes("Machine"))) {
        $machName = $mach.GetAttribute("name")
        if (-not $v2NameSet.ContainsKey($machName)) { continue }

        $info = $v2NameSet[$machName]
        $v3Type = Get-SGridV3TypeName $info.Type  # Get-SGridV3TypeName in SampleGridUpgrade.ps1
        Write-Log "    PARA: $machName -> $v3Type"

        # Update type attribute
        $mach.SetAttribute("type", $v3Type)

        # Remove all existing Parameter children
        $existingParams = @($mach.SelectNodes("Parameter"))
        foreach ($ep in $existingParams) {
            $mach.RemoveChild($ep) | Out-Null
        }

        # Build and add v3 parameters
        $v3Params = Build-SGridV3ParaParams -Xml $xml -V2Type $info.Type

        $globalCount = 0
        $trackCount = 0
        foreach ($p in $v3Params) {
            $mach.AppendChild($p) | Out-Null
            if ($p.GetAttribute("scope") -eq "global") { $globalCount++ }
            else { $trackCount++ }
        }

        $mach.SetAttribute("numGlobalParams", $globalCount.ToString())
        $mach.SetAttribute("numTrackParams", $trackCount.ToString())
        Write-Log "    Set $globalCount global, $trackCount track params"
    }

    # ========================================================================
    # Step 5: Upgrade MACH section - update DLL name, remap state blobs
    # ========================================================================
    if ($machEl) {
        Write-Log "  Upgrading MACH section..."
        foreach ($mach in @($machEl.SelectNodes("Machine"))) {
            $machName = $mach.GetAttribute("name")
            if (-not $v2NameSet.ContainsKey($machName)) { continue }

            $info = $v2NameSet[$machName]
            $v2Dll = $mach.GetAttribute("dll")
            $v3Dll = Get-SGridV3DllName $v2Dll  # Get-SGridV3DllName in SampleGridUpgrade.ps1

            if ($v3Dll) {
                $mach.SetAttribute("dll", $v3Dll)
                Write-Log "    MACH dll: $machName -> $v3Dll"
            } else {
                Write-Log "    WARNING: No DLL mapping for '$v2Dll'"
            }

            # Remap GlobalState - dispatch based on revision
            $gsEl = $mach.SelectSingleNode("GlobalState")
            if ($gsEl -and $gsEl.InnerText) {
                $v2GsBytes = [Convert]::FromBase64String($gsEl.InnerText)
                $v3GsBytes = switch ($info.Revision) {
                    "061110"   { Convert-SGridGlobalState -V2Bytes $v2GsBytes -NumChannels $info.NumChannels }            # Convert-SGridGlobalState in SampleGridUpgrade.ps1
                    "v1switch" { Convert-SGridV1SwitchGlobalState -V1Bytes $v2GsBytes -NumChannels $info.NumChannels }    # Convert-SGridV1SwitchGlobalState in SampleGridUpgrade.ps1
                    "v1byte"   { Convert-SGridV1ByteGlobalState -V1Bytes $v2GsBytes -NumChannels $info.NumChannels }      # Convert-SGridV1ByteGlobalState in SampleGridUpgrade.ps1
                    default    { Convert-SGridOlderGlobalState -V2Bytes $v2GsBytes -NumChannels $info.NumChannels }       # Convert-SGridOlderGlobalState in SampleGridUpgrade.ps1 (older/mid/late share same global layout)
                }
                $gsEl.InnerText = [Convert]::ToBase64String($v3GsBytes)
                Write-Log "    GlobalState ($($info.Revision)): $($v2GsBytes.Length) -> $($v3GsBytes.Length) bytes"
            }

            # Remap TrackState(s) - dispatch based on revision
            foreach ($tsEl in @($mach.SelectNodes("TrackState"))) {
                if ($tsEl.InnerText) {
                    $v2TsBytes = [Convert]::FromBase64String($tsEl.InnerText)
                    $v3TsBytes = switch ($info.Revision) {
                        "061110"   { Convert-SGridTrackState -V2Bytes $v2TsBytes }       # Convert-SGridTrackState in SampleGridUpgrade.ps1
                        "older"    { Convert-SGridOlderTrackState -V2Bytes $v2TsBytes }  # Convert-SGridOlderTrackState in SampleGridUpgrade.ps1
                        "mid"      { Convert-SGridMidTrackState -V2Bytes $v2TsBytes }    # Convert-SGridMidTrackState in SampleGridUpgrade.ps1
                        "late"     { Convert-SGridLateTrackState -V2Bytes $v2TsBytes }   # Convert-SGridLateTrackState in SampleGridUpgrade.ps1
                        { $_ -in @("v1switch", "v1byte") } { Convert-SGridV1TrackState -V1Bytes $v2TsBytes }  # Convert-SGridV1TrackState in SampleGridUpgrade.ps1
                        default    { Convert-SGridOlderTrackState -V2Bytes $v2TsBytes }
                    }
                    $tsEl.InnerText = [Convert]::ToBase64String($v3TsBytes)
                }
            }

            # Convert machine data blob
            # v3 requires a valid data blob - without one the machine crashes on load
            $expectedV3DataSize = 2 + ($info.NumChannels * 5) + ($info.NumChannels * 11) + 5210
            $v3DataBytes = $null
            $dataAttr = $mach.GetAttribute("data")
            if ($dataAttr -and $dataAttr.Length -gt 0) {
                $v2DataBytes = [Convert]::FromBase64String($dataAttr)
                if ($info.Revision -in @("061110", "late", "mid") -and $v2DataBytes.Length -gt 1 -and $v2DataBytes[0] -eq 0x02 -and $v2DataBytes[1] -in @(0x07, 0x08)) {
                    # v2 061110/late (version 7) and mid (version 8) data blobs: MDK prefix + peer data conversion
                    $v3DataBytes = Convert-SGridMachineData -V2Data $v2DataBytes -NumTriggers $info.NumChannels  # Convert-SGridMachineData in SampleGridUpgrade.ps1
                    # Validate converted size matches expected v3 size
                    if ($v3DataBytes.Length -ne $expectedV3DataSize) {
                        Write-Log "    WARNING: Converted data blob size ($($v3DataBytes.Length)) doesn't match expected ($expectedV3DataSize) - using default"
                        $v3DataBytes = $null
                    }
                } else {
                    Write-Log "    WARNING: Data blob ($($v2DataBytes.Length) bytes) from $($info.Revision) cannot be converted"
                    Write-Log "    NOTE: MIDI key assignments and custom group names will be lost"
                    Write-Host "  WARNING: $machName - machine data blob from $($info.Revision) cannot be converted."
                    Write-Host "           MIDI key assignments and custom group names will be reset to defaults."
                }
            } else {
                Write-Log "    No machine data blob to convert"
            }

            # Ensure a valid v3 data blob exists - build default if needed
            if (-not $v3DataBytes) {
                $v3DataBytes = Build-SGridV3DefaultDataBlob -NumChannels $info.NumChannels  # Build-SGridV3DefaultDataBlob in SampleGridUpgrade.ps1
                Write-Log "    Using default v3 data blob ($($v3DataBytes.Length) bytes) for $machName"
            }
            $mach.SetAttribute("data", [Convert]::ToBase64String($v3DataBytes))
            $mach.SetAttribute("dataSize", $v3DataBytes.Length.ToString())
        }
    }

    # ========================================================================
    # Step 6: Upgrade PATT section - remap parameter data rows
    # ========================================================================
    if ($pattEl) {
        Write-Log "  Upgrading PATT section..."
        foreach ($machPatt in @($pattEl.SelectNodes("MachinePatterns"))) {
            $machName = $machPatt.GetAttribute("machine")
            if (-not $v2NameSet.ContainsKey($machName)) { continue }

            $info = $v2NameSet[$machName]
            $numChannels = $info.NumChannels
            $numTracks = [int]$machPatt.GetAttribute("numTracks")

            # Calculate row sizes based on revision
            $v3GlobalRowSize = $numChannels + 17
            $v3TrackRowSize = 30  # 26 params without dividers

            # Source row sizes per revision
            switch ($info.Revision) {
                "061110" {
                    $v2GlobalRowSize = $numChannels + 19  # FW(1)+Div(1)+Trigs(N)+Div(1)+14 common+Inertia(2)
                    $v2TrackRowSize = 35                   # 31 params with 5 divider bytes
                }
                "older" {
                    $v2GlobalRowSize = $numChannels + 18  # FW(1)+Div(1)+Trigs(N)+Div(1)+TrigType..HumanTune(9)+ShufSize..ShufReset(4)+Inertia(2), no LenUnit
                    $v2TrackRowSize = 15                   # 15 params, no dividers
                }
                "mid" {
                    $v2GlobalRowSize = $numChannels + 18  # Same global layout as older
                    $v2TrackRowSize = 23                   # 20 params: Trig(1)+WN(1)+Cmd1(1)+Arg1(1)+Cmd2(1)+Arg2(1)+Subdiv(1)+Mute(1)+Offset(2)+NoteCut(1)+Vol(1)+Vel(1)+HV(1)+Pan(1)+HP(1)+Tune(1)+HT(1)+LF(2)+LFM(2)+Grp(1)
                }
                "late" {
                    $v2GlobalRowSize = $numChannels + 18  # Same global layout as older
                    $v2TrackRowSize = 28                   # 26 params: adds VolEnv/Len, PanEnv/Len, TuneEnv/Len; LpFitMode is byte not word
                }
                "v1switch" {
                    $v2GlobalRowSize = $numChannels + 11  # FW(1)+Trigs(N)+Solo(1)+Vol(1)+Pan(1)+Tune(1)+ShufSize..ShufReset(4)+Inertia(2), NO dividers
                    $v2TrackRowSize = 9                    # WaveNo+Command+Argument+Subdiv+Mute+Vol+Pan+Tune+AuxGroup
                }
                "v1byte" {
                    $v2GlobalRowSize = $numChannels + 14  # FW(1)+Div(1)+Trigs(N)+Div(1)+TrigType(1)+Solo(1)+Vol(1)+Pan(1)+Tune(1)+ShufSize(1)+ShufStep(1)+ShufRnd(1)+ShufReset(1)+Inertia(2) = N+14
                    $v2TrackRowSize = 9                    # Same track params as v1switch
                }
            }

            foreach ($pat in @($machPatt.SelectNodes("Pattern"))) {
                $patLength = [int]$pat.GetAttribute("rows")
                if ($patLength -eq 0) { continue }

                # Find ParamData element
                $pdEl = $pat.SelectSingleNode("ParamData")
                if (-not $pdEl -or -not $pdEl.InnerText) { continue }

                $v2Data = [Convert]::FromBase64String($pdEl.InnerText)
                $expectedSize = ($v2GlobalRowSize + ($v2TrackRowSize * $numTracks)) * $patLength

                if ($v2Data.Length -ne $expectedSize) {
                    Write-Log "    WARNING: PATT data size mismatch for $machName pattern ($($info.Revision)). Expected $expectedSize, got $($v2Data.Length). Skipping."
                    continue
                }

                $commonArgs = @{
                    V2Data = $v2Data
                    V2GlobalRowSize = $v2GlobalRowSize
                    V2TrackRowSize = $v2TrackRowSize
                    V3GlobalRowSize = $v3GlobalRowSize
                    V3TrackRowSize = $v3TrackRowSize
                    NumTracks = $numTracks
                    NumRows = $patLength
                    NumChannels = $numChannels
                }

                $v3Data = switch ($info.Revision) {
                    "061110"   { Convert-SGridPattData @commonArgs }          # Convert-SGridPattData in SampleGridUpgrade.ps1
                    "older"    { Convert-SGridOlderPattData @commonArgs }     # Convert-SGridOlderPattData in SampleGridUpgrade.ps1
                    "mid"      { Convert-SGridMidPattData @commonArgs }       # Convert-SGridMidPattData in SampleGridUpgrade.ps1
                    "late"     { Convert-SGridLatePattData @commonArgs }      # Convert-SGridLatePattData in SampleGridUpgrade.ps1
                    "v1switch" { Convert-SGridV1SwitchPattData @commonArgs }  # Convert-SGridV1SwitchPattData in SampleGridUpgrade.ps1
                    "v1byte"   { Convert-SGridV1BytePattData @commonArgs }    # Convert-SGridV1BytePattData in SampleGridUpgrade.ps1
                    default    { Convert-SGridOlderPattData @commonArgs }
                }

                $pdEl.InnerText = [Convert]::ToBase64String($v3Data)
                Write-Log "    PATT ($($info.Revision)): $machName pattern '$($pat.GetAttribute("name"))' data $($v2Data.Length) -> $($v3Data.Length) bytes"
            }
        }
    }

    # ========================================================================
    # Step 7: Upgrade PAT2 section - rebuild columns in v3 order
    # Instead of remapping old column indices (which preserves wrong column
    # order and misses new v3 params), rebuild the entire column set from
    # scratch using the v3 layout: 24 globals + 26 track params per track.
    # ========================================================================
    if ($pat2El) {
        Write-Log "  Upgrading PAT2 section..."
        foreach ($machPat2 in @($pat2El.SelectNodes("MachineData"))) {
            $machName = $machPat2.GetAttribute("machine")
            if (-not $v2NameSet.ContainsKey($machName)) { continue }

            $info = $v2NameSet[$machName]
            $numGlobalParams = $info.NumChannels + 16  # N triggers + 16 other global params = v3 global count
            $numTrackParams = 26  # v3 always has 26 track params

            # Get numTracks from the PATT section for this machine
            $pattMach = $pattEl.SelectSingleNode("MachinePatterns[@machine='$machName']")
            $numTracks = if ($pattMach) { [int]$pattMach.GetAttribute("numTracks") } else { $info.NumChannels }

            foreach ($pat in @($machPat2.SelectNodes("Pattern"))) {
                $patName = $pat.GetAttribute("name")

                # Remove all existing columns
                $existingCols = @($pat.SelectNodes("Column"))
                foreach ($col in $existingCols) {
                    $pat.RemoveChild($col) | Out-Null
                }

                # Build fresh v3 columns: globals first, then tracks in order
                $colCount = 0

                # Global params (group=1, indexInGroup 0..numGlobalParams-1)
                for ($g = 0; $g -lt $numGlobalParams; $g++) {
                    $col = $xml.CreateElement("Column")
                    $col.SetAttribute("targetMachine", $machName)
                    $col.SetAttribute("group", "1")
                    $col.SetAttribute("indexInGroup", $g.ToString())
                    $col.SetAttribute("track", "0")
                    $col.SetAttribute("numEvents", "0")
                    $pat.AppendChild($col) | Out-Null
                    $colCount++
                }

                # Track params (group=2, for each track: indexInGroup 0..numTrackParams-1)
                for ($t = 0; $t -lt $numTracks; $t++) {
                    for ($tp = 0; $tp -lt $numTrackParams; $tp++) {
                        $col = $xml.CreateElement("Column")
                        $col.SetAttribute("targetMachine", $machName)
                        $col.SetAttribute("group", "2")
                        $col.SetAttribute("indexInGroup", $tp.ToString())
                        $col.SetAttribute("track", $t.ToString())
                        $col.SetAttribute("numEvents", "0")
                        $pat.AppendChild($col) | Out-Null
                        $colCount++
                    }
                }

                $pat.SetAttribute("columnCount", $colCount.ToString())
                Write-Log "    PAT2: $machName/$patName rebuilt with $colCount v3 columns ($numGlobalParams global + $numTrackParams x $numTracks tracks)"
            }
        }
    }

    # ========================================================================
    # Step 7b: Convert Pattern XP editor data for upgraded machines
    # Pattern XP stores pattern data in its machine data blob with flat
    # paramIndex values. These must be remapped from v2 to v3 layout.
    # ========================================================================
    if ($patxEl -and $machEl) {
        Write-Log "  Converting Pattern XP editor data for upgraded machines..."

        # Build map: hidden editor machine name -> parent SGrid machine info
        $editorsToConvert = @{}
        foreach ($me in @($patxEl.SelectNodes("MachineEditor"))) {
            $mname = $me.GetAttribute("machine")
            if (-not $v2NameSet.ContainsKey($mname)) { continue }

            foreach ($pe in @($me.SelectNodes("PatternEditor"))) {
                $editor = $pe.GetAttribute("editor")
                if ($editor -and $editor -ne "builtin") {
                    $editorsToConvert[$editor] = $v2NameSet[$mname]
                }
            }
        }

        # Convert the data blobs for those editor machines
        foreach ($m in @($machEl.SelectNodes("Machine"))) {
            $mname = $m.GetAttribute("name")
            if (-not $editorsToConvert.ContainsKey($mname)) { continue }

            $parentInfo = $editorsToConvert[$mname]
            $dataAttr = $m.GetAttribute("data")
            if (-not $dataAttr -or $dataAttr.Length -eq 0) { continue }

            $peData = [Convert]::FromBase64String($dataAttr)
            Write-Log "    Converting Pattern XP data for '$mname' (editor of $($parentInfo.Name), $($peData.Length) bytes)"

            $convertedData = Convert-PatternXPData `
                -Data $peData `
                -SGridMachineName $parentInfo.Name `
                -NumChannels $parentInfo.NumChannels `
                -Revision $parentInfo.Revision  # Convert-PatternXPData in SampleGridUpgrade.ps1

            $m.SetAttribute("data", [Convert]::ToBase64String($convertedData))
            $m.SetAttribute("dataSize", $convertedData.Length.ToString())
            Write-Log "    Pattern XP data converted for '$mname'"
        }
    }

    # CONN section: no changes needed - connections are by machine name
    # SEQU section: no changes needed - references are by machine name
    # MACX section: no changes needed - references are by machine name

    # ========================================================================
    # Step 8: Save modified XML and re-encode to BMX
    # ========================================================================
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

    Write-Host "Done! $($sgMachines.Count) SampleGrid machine(s) upgraded to v3."
    Write-Host "Output written to: $OutPath"
    Write-Log "SampleGrid upgrade complete. Output: $OutPath"
}
