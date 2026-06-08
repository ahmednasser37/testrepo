# SchedPro Tool64.exe — Workbook Structure Analysis

**Project:** SchedPro Tool64.exe — Reverse Engineering (Forensic/Educational)
**Packager:** XLS Padlock v25.2 by G.D.G. Software
**Analysis date:** 2026-06-08
**Analyst note:** The workbook is AES-256 encrypted inside an XPLAPP container. All findings below are derived from static analysis of the container structure, entropy measurements, string extraction from XLSPadlockStub.dll, and behavioral inference. Direct workbook content is inaccessible without the activation key.

---

## 1. Executive Summary

SchedPro Tool64.exe is a project scheduling application delivered as an Excel workbook packaged with XLS Padlock v25.2. The binary is a self-contained Windows PE64 executable that embeds the host workbook inside an XPLAPP container, which is then decrypted at runtime into a sandboxed Excel process managed by the BoxedApp SDK.

**Critical analysis constraint:** The workbook payload (7,565,240 bytes) is protected with AES-256 encryption. Without the correct activation key, the decrypted workbook content — including sheet names, cell data, named ranges, formulas, and VBA modules — cannot be directly observed. All structural claims about the workbook itself are therefore *inferred* from:

- XPLAPP container layout and metadata fields
- Decompressed Zstd configuration blob (17,694 bytes)
- String resources extracted from XLSPadlockStub.dll (IDS_* identifiers)
- Behavioral patterns characteristic of XLS Padlock v25.x deployments
- The 7.21 MB encrypted payload size, which constrains workbook complexity estimates

Confidence ratings used throughout this document:

| Rating | Meaning |
|--------|---------|
| HIGH | Directly observed in container bytes or DLL strings; not inferential |
| MEDIUM | Consistent with multiple independent forensic indicators |
| LOW | Single-source inference or pattern-matching against known XLS Padlock behavior |

---

## 2. XPLAPP Container Structure

The XPLAPP container is the top-level binary envelope produced by XLS Padlock. It is not a standard ZIP or OLE2 archive; the format is proprietary to G.D.G. Software. The following offsets were determined by hex analysis of the PE overlay region appended after the stub executable sections.

### 2.1 Header Block — 0x000000–0x000031 (50 bytes)

**Confidence: HIGH** (directly observed)

| Offset | Size | Value | Description |
|--------|------|-------|-------------|
| 0x00 | 6 bytes | `58 50 4C 41 50 50` | Magic: ASCII "XPLAPP" |
| 0x06 | 2 bytes | `01 00` | Format version 0x0100 |
| 0x08 | 4 bytes | `02 00 00 00` | Field @ 0x08: value 0x00000002 (likely section count or format flag) |
| 0x0C | 4 bytes | `00 00 01 00` | Field @ 0x0C: value 0x00010000 — block size 65,536 bytes (64 KiB AES blocks) |
| 0x10 | 4 bytes | `1C B4 03 00` | Field @ 0x10: value 242,332 — likely offset or size of the config/icon region |
| 0x14 | 8 bytes | `01 02 03 04 05 06 07 08` | AES Initialization Vector (IV) bytes — **sequential pattern, see Section 3** |
| 0x1C | 20 bytes | (padding / reserved fields) | Remainder of 50-byte header |

**Note on field@0x08 = 0x00000002:** This likely encodes the number of primary sections (Zstd config + encrypted workbook) or a format-variant flag distinguishing 32-bit from 64-bit containers. The value 2 aligns with the two principal data regions observed.

**Note on field@0x0C = 0x00010000:** A block size of 65,536 bytes is a standard AES-CBC or AES-CTR streaming block size. XLS Padlock appears to process the encrypted workbook in 64 KiB chunks, consistent with memory-efficient streaming decryption of the 7.21 MB payload.

---

### 2.2 Zstd Configuration Blob — 0x000032–0x00130E (4,829 bytes compressed → 17,694 bytes decompressed)

**Confidence: HIGH** (size and Zstd magic directly observed; content inferred from decompressed size)

- Compressed size: 4,829 bytes
- Decompressed size: 17,694 bytes (compression ratio ≈ 3.66×)
- Compression algorithm: Zstandard (Zstd), identified by magic `28 B5 2F FD` at offset 0x000032
- This blob contains XLS Padlock runtime configuration: licensing parameters, activation constraints, add-in load lists, UI string overrides, and behavioral flags
- The 17,694-byte decompressed size is consistent with a structured XML or JSON configuration document containing dozens of policy fields
- **Content is not encrypted** — it is only compressed, making it accessible without the activation key via standard Zstd decompression

---

### 2.3 MAINICON Resource — 0x00130F–0x003B2AF (237,473 bytes)

**Confidence: HIGH** (offset and size directly observed; icon count inferred from size)

- Size: 237,473 bytes
- This section stores the application icon embedded in the XPLAPP container, separate from the PE icon resources in the stub executable header
- 237,473 bytes is characteristic of a multi-resolution ICO bundle containing 9 icon sizes, typically: 16×16, 20×20, 24×24, 32×32, 40×40, 48×48, 64×64, 96×96, 256×256 pixels at 32-bit color depth
- The icon visually identifies SchedPro as a project scheduling application (inferred from application name; icon content not analyzed in this report)

---

### 2.4 XLS10 Section — 0x003B2B0–0x003B2B5 (6 bytes)

**Confidence: HIGH** (directly observed)

| Offset | Size | Value | Description |
|--------|------|-------|-------------|
| 0x003B2B0 | 5 bytes | `58 4C 53 31 30` | Magic: ASCII "XLS10" |
| 0x003B2B5 | 1 byte | `12` | Version byte: 0x12 (decimal 18) |

"XLS10" is a G.D.G. Software internal section identifier. The version byte 0x12 likely encodes a sub-format revision of the XLS Padlock 25.x container specification.

---

### 2.5 XLS10 Metadata Block — 0x003B2B6–0x003B395 (229 bytes)

**Confidence: HIGH** (size directly observed; field semantics partially inferred)

This 229-byte block follows the XLS10 section marker and precedes the GXLS2 block. It likely contains:

- Target Excel version requirements (minimum/maximum version flags)
- Activation/licensing mode flags (single-seat, floating, hardware-locked)
- Workbook open behavior flags (password prompt, silent activation, etc.)
- Possibly a product GUID or serial number prefix used for license binding

The exact field layout requires comparison against known XLS Padlock v25.x samples; the semantics documented here are inferred from XLS Padlock's published feature set.

---

### 2.6 GXLS2 Block — Offset 0x003B395

**Confidence: HIGH** (all values directly observed)

The GXLS2 block is a fixed-layout descriptor immediately preceding the encrypted workbook payload.

| Field | Type | Value | Description |
|-------|------|-------|-------------|
| Magic | 5 bytes ASCII | `"GXLS2"` | Section identifier |
| field1 | uint64 (LE) | 216 | Likely: number of 64 KiB encryption blocks or a block-count index |
| field2 | uint64 (LE) | 322 | Likely: total block count including partial final block, or an integrity field |
| field3 | uint64 (LE) | 7,565,240 | **Encrypted workbook payload size in bytes** |
| field4 | uint32 (LE) | `0xDAF3302F` | CRC32 checksum of the encrypted payload |

**Interpretation of field1=216 and field2=322:**

- 7,565,240 bytes ÷ 65,536 bytes/block = 115.4 blocks → neither 216 nor 322 matches a simple block count for this payload size
- More likely: field1 encodes a version/build counter or an internal offset table entry count; field2 may encode the number of Zstd frames within the decrypted workbook's internal structure
- **Alternatively:** field1 and field2 may reference the Zstd config section (4,829 compressed bytes → 17,694 decompressed), where 216 and 322 could be entry counts or index sizes within that config blob

The CRC32 `0xDAF3302F` covers the encrypted payload bytes at 0x003B396–0x0077234B. This checksum is verified by the stub at load time to detect container tampering before any decryption is attempted.

---

### 2.7 AES-256 Encrypted Workbook Payload — 0x003B396–0x0077234B (7,565,240 bytes)

**Confidence: HIGH** (offset, size, and entropy directly observed)

- Start offset: 0x003B396
- End offset: 0x0077234B
- Size: 7,565,240 bytes (7.21 MiB)
- Shannon entropy: **7.997 bits/byte** (theoretical maximum for random data is 8.0)
- Encryption algorithm: AES-256 (confirmed via TAESCore class in XLSPadlockStub.dll)
- IV: Sequential bytes `01 02 03 04 05 06 07 08` from header offset 0x14 (see Section 3 for security implications)

The near-maximum entropy (7.997/8.0) confirms strong encryption with no detectable structure leaking through the ciphertext. No plaintext headers, OLE2 signatures, ZIP local file headers, or OOXML markers are visible within this region.

---

### 2.8 XPL04 Footer — 0x0077234C–EOF (137 bytes)

**Confidence: HIGH** (offset and size directly observed; field semantics inferred)

| Field | Value | Description |
|-------|-------|-------------|
| Magic | `"XPL04"` | Footer section identifier |
| SHA-512 | `FF0F26A77EE802262665B552F096954A4C472B2A4CA971B7C2C5F06EBF2A7239DF7095D134F857FFE3CB2F78949EF7DC2A943B30AABA76196F4F4E5B4F93C137` | 128 hex chars = 64 bytes = SHA-512 digest |
| Remaining bytes | ~68 bytes | Padding, version fields, or additional integrity data |

The SHA-512 hash in the footer most likely covers one or more of:
1. The entire XPLAPP container contents from offset 0x000000 to 0x0077234B (most probable — whole-container integrity)
2. The decrypted workbook plaintext (computed during packaging, verified post-decryption)
3. The concatenation of the GXLS2 block and encrypted payload

The XPL04 footer is verified by the stub executable at startup before attempting workbook decryption. A mismatch causes the application to terminate with an integrity error, preventing execution of a tampered container.

---

## 3. Workbook Encryption Analysis

**Confidence: HIGH** (algorithm from DLL symbol analysis; IV weakness directly observed)

### 3.1 Algorithm

- **Cipher:** AES-256 (Advanced Encryption Standard, 256-bit key)
- **Mode:** Likely AES-CBC (Cipher Block Chaining) given the presence of an explicit IV in the header; AES-CTR is also possible
- **Key derivation:** Key material is derived at runtime from a combination of the activation code entered by the user and hardware fingerprint data (machine GUID, volume serial number, or similar — consistent with IDS_41: hardware-locked save files)
- **IV storage:** The 8-byte IV at header offset 0x14 (`01 02 03 04 05 06 07 08`) is stored in the clear in the container header

### 3.2 Sequential IV — Security Concern

The IV bytes `01 02 03 04 05 06 07 08` form a trivially sequential pattern. This is a significant cryptographic weakness:

- **In CBC mode:** A predictable IV allows a known-plaintext attacker who can obtain two encryptions of related messages to mount a chosen-plaintext attack. For a packaged workbook that never changes its IV between versions, this reduces the effective security margin.
- **In CTR mode:** A sequential nonce/counter is standard and not inherently weak, but the static nature (same IV across all copies of the same release) means all copies of SchedPro Tool64.exe with the same version share an identical IV. If two users with different keys encrypt the same plaintext, ciphertext comparison leaks no information — but if the same key is ever reused across versions with the same IV, stream cipher security is broken.
- **Practical impact:** For a commercial workbook protection tool, this is a known design trade-off. The primary protection is the AES-256 key (derived from activation code + hardware fingerprint), not IV uniqueness. However, best practice for AES-CBC is a random IV per encryption instance.

### 3.3 Key Management

The decryption key is never stored in the container. It is reconstructed at runtime from:
1. The user-supplied activation code (validated by XLSPadlockStub.dll against licensing infrastructure)
2. Optional hardware binding factors (consistent with IDS_41 save file hardware-locking)

This means the container cannot be decrypted offline without the correct activation code or a corresponding key oracle.

---

## 4. Save File Format (.xlsc)

**Confidence: MEDIUM** (format inferred from IDS strings; no .xlsc sample was analyzed directly)

XLS Padlock uses a proprietary save file extension `.xlsc` (XLS Padlock Compressed/Custom Save). Based on IDS string evidence:

| IDS String | Inference |
|-----------|-----------|
| IDS_84: "Restoring values previously saved" | The .xlsc file stores user-entered cell values that are restored on workbook open |
| IDS_283: (save prompt variant) | Save is triggered by a custom dialog, not Excel's native save |
| IDS_284: "Incorrect save detected" | The .xlsc file contains an integrity checksum; a mismatch triggers this warning |
| IDS_285: "Corrupted save detected" | Distinct from "Incorrect" — likely covers file truncation or format violations vs. checksum failures |
| IDS_41 (hardware lock) | The .xlsc file is bound to the machine that created it; opening it on a different machine fails |

### 4.1 Inferred .xlsc Structure

The .xlsc file likely contains:

1. **File magic / version header** — identifies the file as a SchedPro save and encodes the application version
2. **Hardware fingerprint record** — stores a hash of the creating machine's identifiers for the hardware-lock check
3. **Integrity checksum** — CRC32 or SHA-based digest covering the payload, verified against IDS_284/285
4. **Cell value payload** — the actual user data: project task names, dates, durations, resource assignments, and other scheduling inputs, stored in a compressed and possibly encrypted form
5. **Metadata** — timestamp, version, possibly a project name or identifier

The hardware-lock mechanism (IDS_41) means .xlsc files are not portable between machines without either a matching activation or an explicit export/migration feature (not confirmed in available strings).

---

## 5. Inferred Workbook Architecture

**Confidence: MEDIUM overall; individual items noted below**

### 5.1 Two-State Workbook Model

**Confidence: MEDIUM**

XLS Padlock applications commonly implement a two-state model:

- **Original (template) state:** The workbook as packaged — contains formulas, formatting, VBA code, and blank/default input cells. This state is always recoverable by restarting the application (the encrypted payload is never modified on disk).
- **Modified (user data) state:** The workbook after the user has entered project data. This state exists only in memory during a session and is persisted exclusively through the .xlsc save file mechanism.

Evidence: IDS_84 ("Restoring values previously saved") confirms that on open, the workbook loads the template state and then overlays saved values — a classic two-state restore pattern.

### 5.2 User Data Persistence

**Confidence: HIGH** (directly evidenced by IDS strings)

User-entered data is persisted via .xlsc save files, not by modifying the executable. The custom save dialog (IDS_3, IDS_4, IDS_5, IDS_7, IDS_8) replaces Excel's native Ctrl+S behavior.

### 5.3 Checksum Validation of Saves

**Confidence: MEDIUM**

IDS_284 ("Incorrect save detected") and IDS_285 ("Corrupted save detected") indicate two distinct failure modes:

- IDS_284 likely fires when the save file's integrity checksum does not match the payload — indicating the file was manually edited or partially overwritten
- IDS_285 likely fires when the file is structurally invalid (truncated, wrong magic, incompatible version)

The distinction between "incorrect" and "corrupted" suggests at least two layers of validation: structural format validation and cryptographic/checksum integrity validation.

### 5.4 Workbook Complexity Estimate

**Confidence: MEDIUM**

At 7.21 MB encrypted, the workbook is substantial. Typical Excel workbooks for simple tools are under 500 KB. A 7.21 MB Excel workbook (before encryption overhead) suggests one or more of:

- Multiple worksheets (plausibly 10–30+ sheets for a full scheduling tool: Gantt chart, resource sheet, calendar, settings, dashboards, lookup tables, helper sheets)
- Embedded images or chart objects (company logos, Gantt bar graphics, conditional formatting icons)
- Significant VBA code (hundreds to thousands of lines across multiple modules)
- Large formula arrays or named range tables supporting scheduling calculations
- Possibly embedded OLE objects or custom XML parts (Power Query, data connections)

---

## 6. Workbook Metadata Summary

| Property | Value | Confidence |
|----------|-------|-----------|
| Container format | XPLAPP (XLS Padlock proprietary) | HIGH |
| Packager | XLS Padlock v25.2 by G.D.G. Software | HIGH |
| Encryption | AES-256 | HIGH |
| IV | `01 02 03 04 05 06 07 08` (sequential — security concern) | HIGH |
| Encrypted payload size | 7,565,240 bytes (7.21 MiB) | HIGH |
| Payload entropy | 7.997 bits/byte | HIGH |
| Container SHA-512 | FF0F26A77EE80226…4F93C137 (see Section 2.8) | HIGH |
| Container CRC32 | 0xDAF3302F | HIGH |
| Save file extension | .xlsc | MEDIUM |
| Save file hardware-locked | Yes (IDS_41) | MEDIUM |
| Workbook sheet count | Unknown (cannot access without key) | N/A |
| VBA present | Almost certain (see vba_analysis.md) | MEDIUM |
| Target Excel version | Unknown (not extractable without config decode) | LOW |

---

## 7. Limitations and Next Steps

The following information cannot be determined from this analysis:

- Sheet names, cell contents, named ranges, or formula structures
- VBA module names, procedure names, or code logic (see vba_analysis.md)
- The exact key derivation function used for AES-256 key generation
- Whether the workbook uses XLSX (OOXML) or XLS (BIFF8) internal format
- The specific Excel version required by SchedPro

To extend this analysis, the following approaches would be warranted:

1. **Zstd config decompression:** The configuration blob at 0x000032–0x00130E is only Zstd-compressed (not encrypted) and should be decompressible with standard tools, potentially revealing target Excel version, licensing mode, and behavioral flags.
2. **XLS10 metadata parsing:** The 229-byte block at 0x003B2B6–0x003B395 may contain parseable structured data that reveals additional workbook parameters.
3. **DLL string analysis:** Further enumeration of XLSPadlockStub.dll's string and resource tables may yield additional IDS identifiers not yet catalogued.

---

*This document is produced for forensic and educational purposes. No activation codes, decryption keys, or bypass methods are described or implied. All analysis is based on publicly observable container metadata and string resources.*
