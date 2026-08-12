<#PSScriptInfo

.VERSION 1.0.0

.GUID be2c4108-11d3-4ac9-a508-b8d90d20eaba

.AUTHOR Igor Cotruta

.COMPANYNAME

.COPYRIGHT (c) 2026 Igor Cotruta. MIT License.

.TAGS PowerPivot VertiPaq PBIX PBIT Excel xlsx AnalysisServices Tabular xpress8 ABF Windows PSEdition_Desktop PSEdition_Core

.LICENSEURI https://github.com/Hugoberry/PowerPivotPeek/blob/main/LICENSE

.PROJECTURI https://github.com/Hugoberry/PowerPivotPeek

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
1.0.0 - Initial release. Lists and decompresses the embedded Analysis Services
backup files inside a Power Pivot .xlsx (or a legacy uncompressed .pbix), using
ntdll's RtlDecompressBuffer for the xpress8 chunks. No external dependencies.

#>

# Keep the comment-based help block below as the LAST comment before the param
# block. Test-ScriptFileInfo reads .DESCRIPTION via $ast.GetHelpContent(), which
# binds to the final contiguous comment run before the script body -- and
# #Requires tokenizes as a comment. Moving the line below underneath the help
# block hides the description and breaks Publish-Script.
#Requires -Version 5.1

<#
.SYNOPSIS
  Peek into the Power Pivot / Analysis Services data model embedded in an .xlsx
  (or a legacy .pbix), using nothing but what ships with Windows.

.DESCRIPTION
  An Excel workbook that contains a Power Pivot model stores an entire Analysis
  Services backup (ABF) inside the ZIP member 'xl/model/item.data'. Legacy Power
  BI files store the same structure in a member named 'DataModel'.

  This script reads that backup and reports what is inside it:

    * the backup-log header (bytes 72..4096, UTF-16, null-padded XML)
    * the virtual directory, which maps storage paths to offsets and sizes
    * the backup log manifest, which maps storage paths to friendly names
    * a decompressed preview of one embedded file

  The embedded files are compressed with 'xpress8', which is a chunked framing
  around raw MS-XCA Xpress (plain LZ77, no Huffman). Windows can already decode
  that: ntdll's RtlDecompressBuffer with COMPRESSION_FORMAT_XPRESS (3) takes a
  raw MS-XCA buffer, which is exactly what each chunk body is.

  Note that cabinet.dll's Compress/Decompress is the wrong tool here. In buffer
  mode it expects its own container header; the COMPRESS_RAW flag does accept a
  bare buffer, but only in block mode, where you must supply the exact original
  uncompressed size and drive the blocks yourself. RtlDecompressBuffer takes the
  bare buffer directly, which is why this fits in a single P/Invoke.

  This is deliberately a *glimpse*. All the heavy VertiPaq work - column
  segments, RLE runs, dictionaries, hash indexes - is left to pbixray
  (https://github.com/Hugoberry/pbixray).

  Models compressed with XPress9 are detected and reported, but not decoded;
  the Windows compression APIs cannot decode XPress9. Use pbixray for those.

  Windows only: it P/Invokes ntdll. Works in Windows PowerShell 5.1 and in
  PowerShell 7+ on Windows.

.PARAMETER Path
  Path to a .xlsx workbook with an embedded Power Pivot model, or to a legacy
  PBIX file whose DataModel stream is an uncompressed ABF. (Do not start a help
  line with '.pbix' or any other dot-word: the help parser reads it as an
  unknown directive and silently discards this entire comment block.)

.PARAMETER ListFiles
  List every embedded backup file: size, offset and friendly name.

.PARAMETER Dump
  Friendly-path substring of the embedded file to decompress and preview. The
  smallest match wins. Without this, the smallest .xml metadata file is shown.

.PARAMETER Full
  Print the whole decoded file instead of the first 1400 characters.

.PARAMETER OutFile
  Save the full decompressed bytes of the selected file to this path.

.EXAMPLE
  Show-PowerPivotModel -Path Book.xlsx

  Print the backup-log summary and preview the smallest embedded XML file.

.EXAMPLE
  Show-PowerPivotModel -Path Book.xlsx -ListFiles

  List every embedded file in the model, largest first.

.EXAMPLE
  Show-PowerPivotModel -Path Book.xlsx -Dump 'Sales' -OutFile sales.xml

  Decompress the smallest embedded file whose path contains 'Sales' and write
  the full decoded bytes to sales.xml.

.LINK
  https://github.com/Hugoberry/PowerPivotPeek

.LINK
  https://github.com/Hugoberry/pbixray
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Path,
    [switch] $ListFiles,
    [string] $Dump,
    [switch] $Full,
    [string] $OutFile
)

$ErrorActionPreference = 'Stop'

# $IsWindows only exists in PowerShell 6+; Windows PowerShell is always Desktop.
if (-not ($PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows)) {
    throw "Windows only: this script decompresses via ntdll!RtlDecompressBuffer."
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace PP {
  public static class Nt {
    // COMPRESSION_FORMAT_XPRESS = 3 (raw MS-XCA LZ77 -- what xpress8 chunks are)
    [DllImport("ntdll.dll")]
    public static extern int RtlDecompressBuffer(
        ushort fmt, byte[] dst, int dstLen, byte[] src, int srcLen, out int finalSize);
  }
}
'@

# ---------- helpers ----------
function Get-ByteRange([byte[]]$src, [int]$off, [int]$len) {
    if ($off -lt 0 -or $len -lt 0 -or ($off + $len) -gt $src.Length) {
        throw "Range $off..$($off + $len) is outside the $($src.Length)-byte model stream."
    }
    $d = New-Object byte[] $len
    [Array]::Copy($src, $off, $d, 0, $len)
    ,$d
}

function ConvertFrom-XmlBytes([byte[]]$bytes) {
    # Honors the <?xml encoding=...?> declaration (utf-8 or utf-16).
    $doc = New-Object System.Xml.XmlDocument
    $doc.Load((New-Object System.IO.MemoryStream (,$bytes)))
    $doc
}

function Get-NodeText([System.Xml.XmlNode]$node, [string]$name) {
    $n = $node.SelectSingleNode("*[local-name()='$name']")
    if ($n) { $n.InnerText } else { $null }
}

# Decompress a VertiPaq slice: xpress8-chunked when ApplyCompression, else raw.
function Expand-VertiPaqSlice([byte[]]$data, [bool]$compressed) {
    if (-not $compressed) { return ,$data }
    $out = New-Object System.IO.MemoryStream
    $pos = 0
    while ($pos -lt $data.Length) {
        if (($pos + 4) -gt $data.Length) { throw "Truncated xpress8 chunk header at offset $pos." }
        $u = [BitConverter]::ToUInt16($data, $pos)      # uncompressed size
        $c = [BitConverter]::ToUInt16($data, $pos + 2)  # compressed size
        $pos += 4
        if (($pos + $c) -gt $data.Length) { throw "Truncated xpress8 chunk body at offset $pos." }
        if ($u -eq $c) {
            $out.Write($data, $pos, $c)                 # stored verbatim
        } else {
            $body = Get-ByteRange $data $pos $c
            $dst = New-Object byte[] $u
            $final = 0
            $st = [PP.Nt]::RtlDecompressBuffer(3, $dst, $u, $body, $c, [ref]$final)
            if ($st -ne 0) { throw ("RtlDecompressBuffer NTSTATUS 0x{0:x8} at offset {1}" -f $st, ($pos - 4)) }
            $out.Write($dst, 0, $final)
        }
        $pos += $c
    }
    ,$out.ToArray()
}

function Format-Preview([byte[]]$b, [int]$max = 1400) {
    if ($b.Length -eq 0) { return "[empty]" }
    $probe = [Math]::Min(200, $b.Length)
    if ($b.Length -ge 2 -and $b[0] -eq 0xff -and $b[1] -eq 0xfe) {
        $s = [System.Text.Encoding]::Unicode.GetString($b)
    } elseif (-not ($b[0..($probe - 1)] | Where-Object { $_ -lt 9 })) {
        $s = [System.Text.Encoding]::UTF8.GetString($b)
    } else {
        $n = [Math]::Min(64, $b.Length)
        return "[binary] " + (($b[0..($n - 1)] | ForEach-Object { $_.ToString('x2') }) -join ' ')
    }
    if ($s.Length -gt $max) { $s.Substring(0, $max) + "`n... (truncated)" } else { $s }
}

# ---------- read the model member ----------
$modelPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
$zip = [System.IO.Compression.ZipFile]::OpenRead($modelPath)
try {
    $entry = $zip.Entries | Where-Object { $_.FullName -in @('xl/model/item.data', 'DataModel') } | Select-Object -First 1
    if (-not $entry) { throw "No embedded model found (need xl/model/item.data or DataModel)." }
    $member = $entry.FullName
    $ms = New-Object System.IO.MemoryStream
    $es = $entry.Open(); $es.CopyTo($ms); $es.Close()
    $buf = $ms.ToArray()
} finally { $zip.Dispose() }

Write-Host "File   : $modelPath"
Write-Host "Member : $member  ($($buf.Length) bytes)"

# ---------- detect format ----------
if ($buf.Length -lt 4096) { throw "Model stream is only $($buf.Length) bytes; too small to hold an ABF header." }
$head = [System.Text.Encoding]::Unicode.GetString((Get-ByteRange $buf 0 204))
if ($head -like '*STREAM_STORAGE_SIGNATURE*') {
    Write-Host "Format : uncompressed ABF  (winapi xpress8 path applies)`n"
} elseif ($head -like '*XPress9*' -or $head -like '*XPrs9*') {
    Write-Host "Format : XPress9-compressed`n"
    Write-Warning "This model is XPress9-compressed. The Windows compression APIs cannot decode XPress9 -- use pbixray for this file."
    return
} else {
    Write-Warning "Unrecognized model stream signature."
    return
}

# ---------- backup-log header (bytes 72..4096, utf-16, null-padded) ----------
$hdrStr = [System.Text.Encoding]::Unicode.GetString((Get-ByteRange $buf 72 (4096 - 72))).TrimEnd([char]0)
[xml]$hdr = $hdrStr
$applyCompression = (Get-NodeText $hdr.DocumentElement 'ApplyCompression') -eq 'true'
$errorCode        = (Get-NodeText $hdr.DocumentElement 'ErrorCode') -eq 'true'
$filesCount       = [int](Get-NodeText $hdr.DocumentElement 'Files')
$dataSize         = [int](Get-NodeText $hdr.DocumentElement 'DataSize')
$vdOffset         = [int](Get-NodeText $hdr.DocumentElement 'm_cbOffsetHeader')

# ---------- virtual directory (file registry) ----------
$vd = ConvertFrom-XmlBytes (Get-ByteRange $buf $vdOffset $dataSize)
$vdFiles = @{}
$vdList = New-Object System.Collections.Generic.List[object]
foreach ($f in $vd.SelectNodes("//*[local-name()='BackupFile']")) {
    $rec = [pscustomobject]@{
        StoragePath = Get-NodeText $f 'Path'
        Size        = [int](Get-NodeText $f 'Size')
        Offset      = [int](Get-NodeText $f 'm_cbOffsetHeader')
    }
    $vdFiles[$rec.StoragePath] = $rec
    $vdList.Add($rec)
}
if ($vdList.Count -eq 0) { throw "Virtual directory is empty; this is not an ABF layout we understand." }

# ---------- backup-log manifest (last registry entry, uncompressed) ----------
$logEntry = $vdList[-1]
$logLen = if ($errorCode) { $logEntry.Size - 4 } else { $logEntry.Size }
$log = ConvertFrom-XmlBytes (Get-ByteRange $buf $logEntry.Offset $logLen)
$groups = $log.SelectNodes("//*[local-name()='FileGroup']")
$persistRoot = ''
if ($groups.Count -gt 1) { $persistRoot = (Get-NodeText $groups.Item(1) 'PersistLocationPath') + '\' }

# Join friendly paths (backup log) to offsets/sizes (virtual directory).
$files = New-Object System.Collections.Generic.List[object]
foreach ($g in $groups) {
    foreach ($bf in $g.SelectNodes("*[local-name()='FileList']/*[local-name()='BackupFile']")) {
        $sp = Get-NodeText $bf 'StoragePath'
        if ($vdFiles.ContainsKey($sp)) {
            $p = Get-NodeText $bf 'Path'
            if ($persistRoot -and $p.StartsWith($persistRoot)) { $p = $p.Substring($persistRoot.Length) }
            $files.Add([pscustomobject]@{
                Name   = ($p -split '\\')[-1]
                Path   = $p
                Size   = $vdFiles[$sp].Size
                Offset = $vdFiles[$sp].Offset
            })
        }
    }
}

# ---------- summary glimpse ----------
Write-Host "== Backup log =="
Write-Host ("  Object            : {0}" -f (Get-NodeText $log.DocumentElement 'ObjectName'))
Write-Host ("  SyncVersion       : {0}" -f (Get-NodeText $log.DocumentElement 'BackupRestoreSyncVersion'))
Write-Host ("  ApplyCompression  : {0}" -f $applyCompression)
Write-Host ("  ErrorCode trailer : {0}" -f $errorCode)
Write-Host ("  Languages         : {0}" -f (($log.SelectNodes("//*[local-name()='Language']") | ForEach-Object { $_.InnerText }) -join ', '))
Write-Host ("  File groups       : {0}" -f $groups.Count)
Write-Host ("  Embedded files    : {0} (header says {1})" -f $files.Count, $filesCount)
Write-Host ""

if ($ListFiles) {
    Write-Host "== Embedded files (by size) =="
    $files | Sort-Object Size -Descending |
        Format-Table @{N = 'Size'; E = { '{0:N0}' -f $_.Size }; A = 'right' }, Offset, Name -AutoSize |
        Out-Host
}

# ---------- winapi preview of one embedded file ----------
if ($Dump) {
    $target = $files | Where-Object { $_.Path -like "*$Dump*" } | Sort-Object Size | Select-Object -First 1
    if (-not $target) { Write-Warning "No embedded file matching '*$Dump*'." }
} else {
    # default: smallest .xml metadata file, else smallest overall
    $target = $files | Where-Object { $_.Name -like '*.xml' } | Sort-Object Size | Select-Object -First 1
    if (-not $target) { $target = $files | Sort-Object Size | Select-Object -First 1 }
}

if ($target) {
    Write-Host "== winapi preview: $($target.Path) =="
    $sliceLen = if ($errorCode) { $target.Size - 4 } else { $target.Size }
    $raw = Get-ByteRange $buf $target.Offset $sliceLen
    try {
        $dec = Expand-VertiPaqSlice $raw $applyCompression
        Write-Host ("  decompressed {0} -> {1} bytes via ntdll xpress8`n" -f $target.Size, $dec.Length)
        if ($OutFile) {
            $outPath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($OutFile)
            [System.IO.File]::WriteAllBytes($outPath, $dec)
            Write-Host "  saved full file -> $outPath`n"
        }
        Write-Host (Format-Preview $dec $(if ($Full) { [int]::MaxValue } else { 1400 }))
    } catch {
        Write-Warning "decode failed: $_"
    }
}
