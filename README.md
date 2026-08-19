# PowerPivotPeek

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/Show-PowerPivotModel)](https://www.powershellgallery.com/packages/Show-PowerPivotModel)
[![Downloads](https://img.shields.io/powershellgallery/dt/Show-PowerPivotModel)](https://www.powershellgallery.com/packages/Show-PowerPivotModel)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Look inside the Power Pivot data model hidden in an Excel workbook — with nothing
but what already ships with Windows. No Python, no pip, no Cython, no NuGet.

```powershell
Install-Script -Name Show-PowerPivotModel
Show-PowerPivotModel -Path .\Book.xlsx -ListFiles
```

```text
File   : C:\models\Book.xlsx
Member : xl/model/item.data  (4913152 bytes)
Format : uncompressed ABF  (winapi xpress8 path applies)

== Backup log ==
  Object            : Microsoft_SQLServer_AnalysisServices
  SyncVersion       : 1
  ApplyCompression  : True
  ErrorCode trailer : True
  Languages         : 1033
  File groups       : 2
  Embedded files    : 214 (header says 214)

== Embedded files (by size) ==
     Size Offset   Name
     ---- ------   ----
  892,384 4096     Sales_bfd1.Order Date.0.idf
  611,208 896480   Sales_bfd1.Product Key.0.idf
  ...
```

## Why this exists

An `.xlsx` workbook that contains a Power Pivot model is not just a spreadsheet.
Inside the ZIP, the member `xl/model/item.data` holds a complete **Analysis
Services backup file (ABF)** — the same structure a `.pbix` keeps in its
`DataModel` member. Power Pivot has always been a local SSAS Tabular instance in
a trench coat; `msmdsrv.exe` shows up in Process Explorer when you use it.

Getting at that model normally means opening Excel. This script reads it straight
off disk in under 200 lines of PowerShell — and the part that actually does the
decompression is about thirty of them.

You can confirm you are looking at an ABF before doing anything else. The first
72 bytes are a UTF-16 magic string, and it is not subtle:

```text
STREAM_STORAGE_SIGNATURE_)!@#$%^&*(
```

That is 35 characters behind a two-byte byte-order mark, which is why the backup
log header starts at byte 72 rather than at some arbitrary offset.

The format reaches past Excel. Analysis Services writes the same container when
you back up a Tabular database, so an `.abf` from an SSAS instance carries the
same signature, the same virtual directory and the same backup log. Compressed
chunks are the usual case there too, because Microsoft's documentation lists
compressing the backup as the default. Power Pivot workbooks, older `.pbix`
files and SSAS Tabular backups are three wrappers around one format.

## The trick: Windows already has the decompressor

The files inside the backup are compressed with what the VertiPaq world calls
**xpress8**. That name suggests something proprietary, but the framing is simple:

```text
repeated:  [uint16 uncompressed_size][uint16 compressed_size][body]
```

…and when the two sizes are equal, the body is stored verbatim. Each compressed
body is a **raw MS-XCA Xpress buffer** — plain LZ77, no Huffman layer.

Windows can already decode that. `ntdll!RtlDecompressBuffer` with
`COMPRESSION_FORMAT_XPRESS` (`3`) takes a raw MS-XCA buffer and hands back the
plaintext. That is a P/Invoke away from PowerShell:

```powershell
[DllImport("ntdll.dll")]
public static extern int RtlDecompressBuffer(
    ushort fmt, byte[] dst, int dstLen, byte[] src, int srcLen, out int finalSize);
```

**The detour that cost me time:** the obvious candidate is `cabinet.dll`'s
Compression API (`CreateDecompressor` / `Decompress`). In its default buffer
mode it will not read these chunks — it expects data wrapped in its own
container header. You *can* opt into raw with the `COMPRESS_RAW` flag, but that
puts you in block mode, where you must supply the exact original uncompressed
size and manage blocks yourself.

`RtlDecompressBuffer` skips all of it: bare input buffer, output buffer, output
size, done. It is documented as a driver DDI, but `ntdll` exports it and user
mode can call it — which is the whole reason this fits in a P/Invoke.

## How it reads the file

```text
.xlsx (ZIP)
  └── xl/model/item.data          ← Analysis Services backup (ABF)
        ├── bytes 0..72           STREAM_STORAGE_SIGNATURE_)!@#$%^&*(
        │                           BOM + 35 chars of UTF-16 = 72 bytes
        ├── bytes 72..4096        backup-log header, null-padded UTF-16 XML
        │                           ApplyCompression, ErrorCode, Files, DataSize,
        │                           m_cbOffsetHeader → where the directory lives
        ├── virtual directory     storage path → offset + size, per file
        ├── backup log manifest   storage path → friendly path (last dir entry)
        └── N embedded files      xpress8-chunked, at the offsets above
```

Two XML documents have to be joined to get anything useful. The **virtual
directory** knows where each file physically sits but names them by opaque
storage path. The **backup log** knows the friendly names (`Sales_bfd1.Order
Date.0.idf`) but not the offsets. Join them on the storage path and you have a
browsable file table.

The `ErrorCode` flag in the header means every embedded file carries a 4-byte
trailer that is not part of its content — miss that and every decode is off by
four bytes at the tail.

Friendly paths are stored absolute, rooted at whatever folder the authoring
machine used, so they need trimming before they are readable. The root to trim
is the shortest `PersistLocationPath` across the file groups, since every group
persists somewhere underneath the database folder. Reaching for a fixed group
index instead lands you on a cube or dimension subfolder and leaves most of the
paths as full temp paths from a stranger's machine.

## Usage

```powershell
# Summary + preview of the smallest XML metadata file
Show-PowerPivotModel -Path .\Book.xlsx

# Every embedded file, largest first
Show-PowerPivotModel -Path .\Book.xlsx -ListFiles

# Decompress a specific file and print all of it
Show-PowerPivotModel -Path .\Book.xlsx -Dump 'cub.xml' -Full

# Decompress and save the raw decoded bytes
Show-PowerPivotModel -Path .\Book.xlsx -Dump 'Sales' -OutFile .\sales.idf
```

| Parameter | Meaning |
| --- | --- |
| `-Path` | `.xlsx` with a Power Pivot model, or a legacy `.pbix` with an uncompressed `DataModel` |
| `-ListFiles` | Print the whole embedded file table, sorted by size |
| `-Dump <substring>` | Pick the file to decode by friendly-path substring (smallest match wins) |
| `-Full` | Print the entire decoded file instead of the first 1400 characters |
| `-OutFile <path>` | Write the full decoded bytes to disk |

## Requirements

- **Windows.** The script P/Invokes `ntdll`. It refuses to run elsewhere.
- Windows PowerShell 5.1 or PowerShell 7+.
- Nothing else.

## Limitations

This is deliberately a *glimpse*, not a parser.

- **XPress9 models are not supported.** Newer and larger models wrap the whole
  backup in XPress9, which the Windows compression APIs cannot decode. The script
  detects this and tells you. Use [pbixray](https://github.com/Hugoberry/pbixray),
  which carries a Cython XPress9 port.
- **No VertiPaq decoding.** You get the `.idf` column segments as bytes. Turning
  those back into rows means dictionaries, hash indexes and RLE runs — again,
  that is what pbixray is for.
- **It prints, it does not pipe.** The script writes a human-readable report via
  `Write-Host`. A `PowerPivotPeek` module exposing proper objects is the planned
  next step. (`PSAvoidUsingWriteHost` is a deliberate, documented exception here:
  the entire purpose of the script is console display.)

If you want tables, relationships, DAX measures and reconstructed row data as
DataFrames, you want **[pbixray](https://github.com/Hugoberry/pbixray)**. This
script is just a small window into the same container.

## Install without the Gallery

It is one self-contained file. Downloading it and running it works fine:

```powershell
irm https://raw.githubusercontent.com/Hugoberry/PowerPivotPeek/main/Show-PowerPivotModel.ps1 -OutFile Show-PowerPivotModel.ps1
.\Show-PowerPivotModel.ps1 -Path .\Book.xlsx
```

## Roadmap

- [ ] `PowerPivotPeek` module exporting `Show-PowerPivotModel`, `Get-PowerPivotFile`, `Export-PowerPivotFile`
- [ ] Pester tests against a committed sample workbook
- [ ] Emit objects to the pipeline instead of formatted text

## Further reading

- [[MS-XCA]: Xpress Compression Algorithm](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-xca/a8b7cb0a-92a6-4187-a23b-5e14273b96f8)
- [[MS-XLDM]: Spreadsheet Data Model File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xldm/8c62e8ce-f605-488d-81e9-4ecdb7686a52)
- [`RtlDecompressBuffer`](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntifs/nf-ntifs-rtldecompressbuffer)
- [Backup and Restore of Analysis Services Databases](https://learn.microsoft.com/en-us/analysis-services/multidimensional-models/backup-and-restore-of-analysis-services-databases) on the `.abf` files that share this container
- [pbixray](https://github.com/Hugoberry/pbixray) — the full Python parser
- [Windows Already Ships the Decompressor Power Pivot Needs](https://www.pbixray.com/posts/powerpivot-xpress8-powershell/) — the write-up behind this script
- [Parsing Power Pivot Data Models from Excel XLSX Files](https://www.pbixray.com/posts/parsing-power-pivot-xlsx/)

## License

MIT — see [LICENSE](LICENSE).
