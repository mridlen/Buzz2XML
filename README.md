# Buzz2XML

A PowerShell tool for converting [Jeskola Buzz](https://jeskola.net/buzz/) song files (.bmx) to and from XML, with support for VST plugin path remapping.

## Features

- **Decode** - Convert .bmx binary files to human-readable XML
- **Encode** - Convert XML back to .bmx with byte-for-byte round-trip fidelity
- **Remap** - Find and replace VST plugin paths embedded in machine data
- **Machines** - List and delete machines from song files
- **GUI** - WinForms frontend for non-technical users

All BMX sections are fully supported: BVER, PARA, MACH, CONN, CONX, MACX, WAVT, PATT, PAT2, PATX, SEQU, BLAH, PDLG, MIDI, WAVE, BGUI.

## Requirements

- Windows PowerShell 5.1+ or PowerShell 7+
- .NET Framework (for WinForms GUI)

## CLI Usage

### Decode (BMX to XML)

```powershell
.\Buzz2XML.ps1 -Mode decode -InputFile mysong.bmx -OutputFile mysong.xml
```

Converts a Buzz .bmx song file into XML. Machine-specific plugin data and some sections are stored as base64 for round-trip fidelity.

### Encode (XML to BMX)

```powershell
.\Buzz2XML.ps1 -Mode encode -InputFile mysong.xml -OutputFile mysong_copy.bmx
```

Converts an XML file (previously created by decode) back into a .bmx binary file. Produces a byte-for-byte identical copy if the XML has not been modified.

### Remap VST Paths

List all embedded VST/DLL paths in a song:

```powershell
.\Buzz2XML.ps1 -Mode remap -InputFile mysong.bmx -ListPaths
```

Replace a path prefix across all machines:

```powershell
.\Buzz2XML.ps1 -Mode remap -InputFile mysong.bmx -OutputFile fixed.bmx `
    -RemapFrom "C:\Program Files (x86)\Jeskola\Buzz\Gear\Vst" `
    -RemapTo "D:\Audio\VST Plugins"
```

### List Machines

```powershell
.\Buzz2XML.ps1 -Mode machines -InputFile mysong.bmx -ListMachines
```

Shows all machines in the song with their type (generator/effect/hidden).

### Delete Machines

Delete by wildcard pattern:

```powershell
.\Buzz2XML.ps1 -Mode machines -InputFile mysong.bmx -OutputFile cleaned.bmx `
    -DeletePattern "SVerb*"
```

Delete specific machines by exact name:

```powershell
.\Buzz2XML.ps1 -Mode machines -InputFile mysong.bmx -OutputFile cleaned.bmx `
    -DeleteNames "SVerb","SVerb2","SVerb22"
```

Both `-DeletePattern` and `-DeleteNames` can be combined. The Master machine cannot be deleted. All references (connections, patterns, sequences, MIDI bindings, etc.) are cleaned up automatically.

### Help

```powershell
.\Buzz2XML.ps1 -Help
```

Also accepts `-h`, `--help`, `/?`, `/h`, `/help`.

## GUI Usage

```powershell
.\Buzz2XML-GUI.ps1
```

Launches a WinForms interface with tabs for Decode, Encode, and Remap operations. Includes file browse dialogs and auto-fills output filenames.

## Notes

- A `.log` file is created alongside the output file for diagnostics.
- Round-trip fidelity has been verified on files with 300+ machines and 1 MB+ pattern sections.
- The remap feature only changes path prefixes, preserving filenames and null-padding the buffer to maintain binary compatibility.
- Machine deletion removes the machine from all sections: PARA, MACH, CONN, PATT, PAT2, PATX, SEQU, MACX, MIDI. Associated hidden pattern editors (pe machines) are auto-detected and removed. CONX and PDLG are dropped since they use opaque machine indices.

## License

MIT
