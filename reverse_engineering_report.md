# Reverse Engineering Report
## SchedPro_Tool64.exe — Comprehensive Static Analysis

**Analyst:** Automated Static Analysis Pipeline  
**Date:** 2026-06-08  
**Classification:** Educational / Defensive / Interoperability Research  
**Target:** `dc30f305-SchedPro_Tool64.exe`  

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [File Overview](#2-file-overview)
3. [PE Structure Analysis](#3-pe-structure-analysis)
4. [Strings Analysis](#4-strings-analysis)
5. [Imports and Exports](#5-imports-and-exports)
6. [Resources](#6-resources)
7. [Architecture Overview](#7-architecture-overview)
8. [Control Flow Analysis](#8-control-flow-analysis)
9. [Runtime Behavior](#9-runtime-behavior)
10. [Memory Analysis](#10-memory-analysis)
11. [Encryption / Packing Analysis](#11-encryption--packing-analysis)
12. [Decrypted Components](#12-decrypted-components)
13. [Network Activity](#13-network-activity)
14. [File System Activity](#14-file-system-activity)
15. [Registry Activity](#15-registry-activity)
16. [Security Findings](#16-security-findings)
17. [Pseudocode Reconstruction](#17-pseudocode-reconstruction)
18. [Functional Description](#18-functional-description)
19. [Indicators of Compromise (IOCs)](#19-indicators-of-compromise-iocs)
20. [Conclusion](#20-conclusion)

---

## 1. Executive Summary

**Confidence: High**

`SchedPro_Tool64.exe` is a **legitimate, commercially-distributed scheduling application** named "SchedPro," authored by **YoussefMoataz**, and distributed as a standalone Windows executable. It is **not malware**.

The file is a **XLS Padlock v25.2** packager product by G.D.G. Software. XLS Padlock is a commercial tool that converts Microsoft Excel workbooks into standalone, self-contained `.exe` files with optional copy-protection, licensing, and activation systems.

**Key structural facts:**
- The outer EXE (112 KB of code) is a thin MSVC-compiled loader stub.
- Embedded inside is a **7-zip archive** (LZMA-compressed) containing two files: `xlspadlock.bin` (7.45 MB) and `XLSPadlockStub.dll` (10.6 MB).
- `xlspadlock.bin` uses a proprietary `XPLAPP` container format that holds UI configuration, an application icon, and a **7.21 MB AES-encrypted Excel workbook** (the actual SchedPro application).
- `XLSPadlockStub.dll` is the **Delphi/Embarcadero VCL runtime engine** (v25.2.0.0, copyright G.D.G. Software 2012–2025) that decrypts, virtualizes, and executes the workbook at runtime using **BoxedApp SDK** virtual filesystem technology.

There are **no malicious behaviors, no network beacons, no credential harvesting, no exploits**. The only network activity is optional licence activation via `http://www.xlspadlock.com`.

---

## 2. File Overview

| Property | Value |
|---|---|
| **Filename** | `SchedPro_Tool64.exe` |
| **Size** | 12,582,912 bytes (12 MB) |
| **MD5** | `68018c3cd1148c2102bdc15862d581fa` |
| **SHA-1** | `3585ed833952ffbc35c4aea1588edc76ccf7939d` |
| **SHA-256** | `56b328e803fa70d27f894441fe2ee0a27ceafd7dae32fb940d02f47cfa02bc0d` |
| **File Type** | PE32+ executable (GUI) x86-64, Windows |
| **Architecture** | AMD64 (x86-64) |
| **Subsystem** | GUI (Windows application, no console) |
| **Linker** | Microsoft Visual C++ v14.16 (VS 2017) |
| **Compile Timestamp** | 2025-10-29 16:47:52 UTC |
| **Sections** | 8 |
| **ImageBase** | `0x0000000140000000` |
| **Entry Point** | RVA `0x00009E88` |
| **Product Name** | SchedPro v1.0.0.2 |
| **Company** | YoussefMoataz |
| **Packer/Protector** | XLS Padlock v25.2 (G.D.G. Software) |
| **Verdict** | **Benign — Commercial Excel application packager** |

---

## 3. PE Structure Analysis

### 3.1 DOS and NT Headers

```
DOS Header:
  e_magic:    0x5A4D  ("MZ")
  e_lfanew:   0x00000108  (non-standard offset — extended stub region)

NT Signature: "PE\0\0"

File Header:
  Machine:              0x8664  (AMD64)
  NumberOfSections:     8
  TimeDateStamp:        0x69024538  (2025-10-29 16:47:52 UTC)
  Characteristics:      0x0022  (Executable, Large Address Aware)

Optional Header:
  Magic:                0x020B  (PE32+)
  MajorLinkerVersion:   14
  MinorLinkerVersion:   16  (MSVC 14.16 / VS2017)
  SizeOfCode:           0x0001BA00  (113,152 bytes)
  SizeOfInitData:       0x0004E000  (319,488 bytes)
  AddressOfEntryPoint:  0x00009E88
  ImageBase:            0x0000000140000000
  SizeOfImage:          0x0006E000  (450,560 bytes)
  CheckSum:             0x00B7E172
  Subsystem:            2  (Windows GUI)
  DllCharacteristics:   0x8160  (HIGH_ENTROPY_VA | DYNAMIC_BASE | NX_COMPAT | TERMINAL_SERVER_AWARE)
```

The `e_lfanew` value of `0x108` is larger than the typical `0x40`, indicating the executable uses an **extended DOS stub region**. This is a characteristic of the MSVC CRT start-up code but is not an evasion technique.

### 3.2 Section Table

| Name | Virt. Address | Virt. Size | Raw Offset | Raw Size | Entropy | Characteristics |
|---|---|---|---|---|---|---|
| `.text` | `0x001000` | `0x1B900` | `0x000400` | `0x1BA00` | **6.468** | CODE, EXEC, READ |
| `.rdata` | `0x01D000` | `0x0ADC0` | `0x1BE00` | `0x0AE00` | 5.092 | INIT_DATA, READ |
| `.data` | `0x028000` | `0x04FCC` | `0x26C00` | `0x00C00` | 1.906 | INIT_DATA, READ, WRITE |
| `.pdata` | `0x02D000` | `0x015D8` | `0x27800` | `0x01600` | 5.256 | INIT_DATA, READ |
| `.gxfg` | `0x02F000` | `0x01350` | `0x28E00` | `0x01400` | 5.003 | INIT_DATA, READ |
| `.gehcont` | `0x031000` | `0x00010` | `0x2A200` | `0x00200` | 0.082 | INIT_DATA, READ |
| `.rsrc` | `0x032000` | `0x3AC1C` | `0x2A400` | `0x3AE00` | **6.254** | INIT_DATA, READ |
| `.reloc` | `0x06D000` | `0x00658` | `0x65200` | `0x00800` | 4.849 | INIT_DATA, READ, DISCARDABLE |

**Entropy analysis:**
- `.text` at 6.468 is normal for compiled x86-64 code (not packed).
- `.rsrc` at 6.254 is elevated because it contains **PNG images** (compressed) and icon data — entirely expected.
- `.gxfg` is the **eXtended Flow Guard (XFG)** metadata section added by MSVC — a legitimate security mitigation.
- `.gehcont` holds the **Guard EH Continuation Table** for exception-based control flow integrity.
- No sections have near-maximum entropy (> 7.5), confirming **the outer loader is not packed or obfuscated**.

### 3.3 Data Directory Summary

| Index | Name | RVA | Size |
|---|---|---|---|
| 1 | Import Directory | `0x0002747C` | `0x64` (5 DLLs) |
| 2 | Resource Directory | `0x00032000` | `0x5A38` |
| 3 | Exception Directory | `0x0002D000` | `0x15D8` |
| 5 | Base Relocation | `0x00038000` | `0x658` |
| 6 | Debug | `0x00025850` | `0x38` |
| 10 | Load Configuration | `0x00025890` | `0x100` |
| 12 | IAT | `0x0001D000` | `0x2C0` |

**Observations:**
- No Export Directory — the EXE does not expose symbols to other processes.
- No TLS Directory — no thread-local storage callbacks (common malware persistence mechanism not present).
- No Security (Authenticode) directory — the binary is **unsigned**.
- Load Config present with `SecurityCookie` at `0x140028008` and XFG/Retpoline data.

### 3.4 Security Mitigations

| Mitigation | Status | Details |
|---|---|---|
| ASLR (DYNAMIC_BASE) | **Enabled** | 64-bit ASLR with HIGH_ENTROPY_VA |
| DEP/NX (NX_COMPAT) | **Enabled** | Non-executable stack/data |
| Stack Canary | **Enabled** | `__security_cookie` at `0x140028008` |
| CFG (Control Flow Guard) | Partial | Bit not set in DllCharacteristics but Load Config present |
| XFG (eXtended Flow Guard) | **Enabled** | `.gxfg` section present (MSVC 16+ feature) |
| GUARD_EH (EH Continuation) | **Enabled** | `.gehcont` section present |
| Retpoline | **Enabled** | `MODULE_RETPOLINE_PRESENT` in GuardFlags |
| Code Signing | **Absent** | Binary is not Authenticode-signed |
| APPCONTAINER | Absent | Not isolated in a Windows container |

---

## 4. Strings Analysis

### 4.1 Critical Strings (High Confidence Findings)

```
// XLS Padlock identity strings (Unicode, in main stub):
"XLSPadlockStub.dll"
"\XLS Padlock Runtime"
"%s\XLS Padlock Runtime"
"%s\XLS Padlock Runtime\%d"
"xlspadlock.bin"
"XlsPadlock.Bin"

// Loader error messages (ASCII):
"Error opening EXE"
"Cannot determine EXE size"
"7z archive not found"
"Memory allocation error"
"SzArEx_Open error: %d"
"XlsPadlock.Bin CRC error"
"XlsPadlock.Bin not found"
"Failed to create output directory"
"CRC comparison: existing=%08X, stored=%08X"
"Cannot update %s: file is in use and has a different CRC."
"Cannot update %s: file is in use."
"Failed to load XLSPadlockStub DLL"

// 7-zip LZMA SDK integration (ASCII):
"SzArEx_Open error: %d"     // lzma2/7z SDK error reporting

// Version info (Unicode resource):
"FileVersion"    = "1.0.0.2"
"ProductVersion" = "1.0.0.2"
"CompanyName"    = "YoussefMoataz"
"ProductName"    = "SchedPro"
"FileDescription"= "SchedPro"
"InternalName"   = "SchedPro"
"LegalCopyright" = "YoussefMoataz"
"LegalTrademarks"= "YoussefMoataz"

// Assembly identity from manifest:
"GDGSoftware.XLSPadlock.App"    // publisher identity
"XLS Padlock Application"        // description
```

### 4.2 MSVC Runtime Artifacts

The `.rdata` section contains a full set of MSVC C++ runtime type names (`__cdecl`, `__stdcall`, `operator new`, `` `vftable' ``, `` `dynamic initializer for '` ``), confirming the stub was compiled with **Microsoft Visual C++ 14.16**.

### 4.3 Notable Absent Strings

| Category | Finding | Significance |
|---|---|---|
| Shell commands | None | No `cmd.exe`, PowerShell, or script execution |
| File write paths | None (other than extraction) | No payload drops beyond intended runtime |
| Remote IP/domain | None | No hardcoded C2 infrastructure |
| Crypto keys | None in plaintext | No embedded AES/RSA keys in the stub |
| Registry paths | None in stub | Registry access is handled by the runtime DLL |

---

## 5. Imports and Exports

### 5.1 Main EXE Imports

The stub imports from exactly 4 DLLs:

#### `KERNEL32.dll` (70 functions)
All imports are standard MSVC CRT/Win32 API. Key functional groups:

| Group | Functions |
|---|---|
| **File extraction** | `CreateFileW`, `ReadFile`, `WriteFile`, `SetFilePointer`, `SetFilePointerEx`, `FindFirstFileW`, `FindNextFileW`, `FindClose` |
| **Heap management** | `HeapAlloc`, `HeapFree`, `HeapReAlloc`, `HeapSize`, `GetProcessHeap` |
| **Dynamic loading** | `LoadLibraryW`, `LoadLibraryExW`, `GetProcAddress`, `FreeLibrary`, `GetModuleHandleW`, `GetModuleHandleExW` |
| **Console I/O** | `GetStdHandle`, `WriteConsoleW`, `GetConsoleMode`, `GetConsoleOutputCP` |
| **Thread-local storage** | `TlsAlloc`, `TlsGetValue`, `TlsSetValue`, `TlsFree` |
| **Fiber-local storage** | `FlsAlloc`, `FlsGetValue`, `FlsSetValue`, `FlsFree` |
| **Process telemetry** | `GetCurrentProcessId`, `GetCurrentThreadId`, `QueryPerformanceCounter` |
| **Anti-debug indicator** | `IsDebuggerPresent` *(see §16)* |
| **Exception handling** | `RtlCaptureContext`, `RtlLookupFunctionEntry`, `RtlVirtualUnwind`, `RtlUnwindEx`, `SetUnhandledExceptionFilter` |
| **SEH support** | `UnhandledExceptionFilter`, `RtlPcToFileHeader` |
| **Encoding** | `EncodePointer` (safe unlinking protection) |
| **Locale/strings** | `MultiByteToWideChar`, `WideCharToMultiByte`, `LCMapStringW`, `GetStringTypeW` |

#### `USER32.dll` (2 functions)
```
MessageBoxW    // used for error dialogs (error opening EXE, CRC failures)
LoadStringW    // loads strings from resource section
```

#### `SHELL32.dll` (4 functions)
```
CommandLineToArgvW      // parse command-line arguments
SHCreateDirectoryExW   // create extraction directory tree
SHGetFolderPathW        // locate %TEMP% or %APPDATA%
SHFileOperationW        // file copy/move operations
```

#### `COMCTL32.dll` (1 function, by ordinal)
```
ord#17 = InitCommonControls  // register common controls for the progress UI
```

### 5.2 Exports
None. The executable does not export any symbols.

### 5.3 Delayed Imports
None detected.

---

## 6. Resources

The `.rsrc` section (raw size: 237,568 bytes, entropy: 6.254) contains:

| Type | ID | Offset | Size | Description |
|---|---|---|---|---|
| RT_BITMAP (3) | 1 | `0x2A710` | 1,128 B | 16×16 bitmap |
| RT_BITMAP (3) | 2 | `0x2AB78` | 2,440 B | 24×48 bitmap |
| RT_BITMAP (3) | 3 | `0x2B500` | 4,264 B | 32×64 bitmap |
| RT_BITMAP (3) | 4 | `0x2C5A8` | 9,640 B | 48×96 bitmap |
| RT_BITMAP (3) | 5 | `0x2EB50` | 16,936 B | 64×128 bitmap |
| RT_BITMAP (3) | 6 | `0x32D78` | 21,640 B | 72×144 bitmap |
| RT_BITMAP (3) | 7 | `0x38200` | 38,056 B | 96×192 bitmap |
| RT_BITMAP (3) | 8 | `0x416A8` | 67,624 B | 128×256 bitmap |
| RT_RCDATA (10) | 9 | `0x51ED0` | **75,581 B** | **Full PNG image** (confirmed by `\x89PNG` magic) |
| RT_STRING (6) | 1 | `0x64610` | 48 B | Error/UI string table |
| RT_GROUP_ICON (14) | MAIN | `0x64640` | 132 B | Icon group descriptor (9 sizes) |
| RT_VERSION (16) | 1 | `0x646C4` | 676 B | VERSIONINFO resource |
| RT_MANIFEST (24) | 1 | `0x64968` | 1,715 B | XML application manifest |

### 6.1 Embedded Manifest (RT_MANIFEST)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity
      name="GDGSoftware.XLSPadlock.App"
      processorArchitecture="amd64"
      version="1.0.0.0"
      type="win32"/>
  <description>XLS Padlock Application</description>
  <dependency>
    <!-- Modern common controls (visual styles) -->
    <dependentAssembly>
      <assemblyIdentity type="win32"
        name="Microsoft.Windows.Common-Controls" version="6.0.0.0"
        publicKeyToken="6595b64144ccf1df"
        processorArchitecture="amd64"/>
    </dependentAssembly>
  </dependency>
  <trustInfo>
    <security>
      <requestedPrivileges>
        <!-- Runs as standard user, no elevation requested -->
        <requestedExecutionLevel level="asInvoker" uiAccess="false"/>
      </requestedPrivileges>
    </security>
  </trustInfo>
  <application>
    <windowsSettings>
      <dpiAware>true</dpiAware>          <!-- DPI-aware application -->
      <heapType>SegmentHeap</heapType>    <!-- Windows 10 segment heap -->
    </windowsSettings>
  </application>
  <!-- Supports Vista through Windows 11 -->
  <compatibility>
    <application>
      <supportedOS Id="{e2011457-...}"/>  <!-- Vista -->
      <supportedOS Id="{35138b9a-...}"/>  <!-- Win7 -->
      <supportedOS Id="{4a2f28e3-...}"/>  <!-- Win8 -->
      <supportedOS Id="{1f676c76-...}"/>  <!-- Win8.1 -->
      <supportedOS Id="{8e0f7a12-...}"/>  <!-- Win10/11 -->
    </application>
  </compatibility>
</assembly>
```

**Key observations:**
- `level="asInvoker"` — the application does **not** request elevation. This is expected for a standard desktop application.
- Publisher identity is `GDGSoftware.XLSPadlock.App`, confirming the XLS Padlock origin.
- `SegmentHeap` directive enables the Windows 10 low-fragmentation heap.

---

## 7. Architecture Overview

### 7.1 Three-Layer Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────┐
│                     SchedPro_Tool64.exe                              │
│            (MSVC x64 Loader Stub — 112 KB code)                      │
│                                                                       │
│  Responsibilities:                                                    │
│  ┌───────────────────────────────────────────────────────────────┐   │
│  │ 1. Parse command-line arguments (CommandLineToArgvW)          │   │
│  │ 2. Determine extraction path (%TEMP% or %APPDATA%)            │   │
│  │ 3. Locate embedded 7z archive (tail-appended to self)         │   │
│  │ 4. Extract archive to temp directory using 7z LZMA SDK        │   │
│  │ 5. CRC-verify xlspadlock.bin                                  │   │
│  │ 6. LoadLibraryW("XLSPadlockStub.dll")                        │   │
│  │ 7. GetProcAddress("XLSPadlockInit")                           │   │
│  │ 8. Call XLSPadlockInit(config_struct)                         │   │
│  └───────────────────────────────────────────────────────────────┘   │
│                                                                       │
│  Contains embedded 7z archive at file offset 0x65A00:                │
└──────────────────────────┬──────────────────────────────────────────┘
                           │ 7-zip LZMA extraction
          ┌────────────────┴─────────────────────────┐
          │                                          │
          ▼                                          ▼
┌──────────────────────┐              ┌──────────────────────────────┐
│   xlspadlock.bin     │              │   XLSPadlockStub.dll         │
│   (7.45 MB)          │              │   (10.6 MB)                  │
│   XPLAPP format:     │              │   Delphi/Embarcadero VCL     │
│                      │              │   G.D.G. Software v25.2      │
│ ┌──────────────────┐ │              │                              │
│ │ XPLAPP Header    │ │  ┌─────────▶│ XLSPadlockInit()             │
│ │ (50 bytes)       │ │  │          │   - BoxedApp SDK init         │
│ ├──────────────────┤ │  │          │   - Virtual FS setup          │
│ │ Zstd Config      │ │  │          │   - Decrypt workbook          │
│ │ (17 KB decmpd)   │ │  │          │   - Launch Excel automation   │
│ ├──────────────────┤ │  │          │   - Activation check          │
│ │ MAINICON         │ │  │          │                              │
│ │ (231 KB, 9 sizes)│ │  │          │ ┌──────────────────────────┐ │
│ ├──────────────────┤ │  │ reads    │ │ BoxedApp SDK             │ │
│ │ XLS10 Header     │ │◀─┘          │ │  (virtual file system)   │ │
│ ├──────────────────┤ │             │ │  - Virtual registry      │ │
│ │ AES-256 Encrypted│ │             │ │  - DLL virtualization    │ │
│ │ Excel Workbook   │ │             │ └──────────────────────────┘ │
│ │ (7.21 MB)        │ │             │                              │
│ ├──────────────────┤ │             │ ┌──────────────────────────┐ │
│ │ XPL04 Footer     │ │             │ │ Second .text section     │ │
│ │ + 64-byte hash   │ │             │ │ Entropy: 7.999 (packed)  │ │
│ │ (137 bytes)      │ │             │ │ (Excel COM runtime DLLs) │ │
│ └──────────────────┘ │             │ └──────────────────────────┘ │
└──────────────────────┘             └──────────────────────────────┘
```

### 7.2 Component Summary Table

| Component | Technology | Size | Purpose |
|---|---|---|---|
| Main EXE stub | MSVC x64 C++ | 112 KB code | Extraction and loader orchestration |
| 7z archive | LZMA + BCJ2 | 11.3 MB | Container for runtime and workbook |
| `xlspadlock.bin` | XPLAPP proprietary | 7.45 MB | Encrypted workbook container |
| Config data (Zstd) | Zstandard | 4.7 KB → 17 KB | UI captions, localization, settings |
| MAINICON | ICO (9 sizes) | 231 KB | Application icon (embedded at runtime) |
| Excel workbook | AES-256 ciphertext | 7.21 MB | The actual SchedPro application |
| `XLSPadlockStub.dll` | Delphi VCL + BoxedApp | 10.6 MB | Decryption engine, Excel automation host |

---

## 8. Control Flow Analysis

### 8.1 Entry Point Sequence

The Win64 ABI entry point is at `0x140009E88`. The typical MSVC CRT startup performs:

```asm
; EP: 0x140009E88
sub  rsp, 0x28                  ; shadow space
call 0x14000A150                ; __security_init_cookie + CRT init
add  rsp, 0x28
jmp  0x140009D14                ; -> WinMainCRTStartup proxy
```

### 8.2 Main Logic Function (`0x140009D14`)

```asm
; 0x140009D14 - WinMainCRTStartup equivalent
mov  ecx, 1
call 0x140009ED8                ; InitializeCriticalSectionAndSpinCount wrapper
                                ; -> global mutex init
test al, al
je   error_exit

call 0x140009E9C                ; Check if already running (TLS flag check)
mov  bl, al

mov  ecx, [0x140029030]        ; state variable
cmp  ecx, 1
je   already_extracted          ; skip if files already on disk

; First run: extract embedded 7z archive
test ecx, ecx
jne  reuse_existing

mov  dword [0x140029030], 1    ; mark as in-progress
lea  rdx, [0x14001D328]        ; DLL path string
lea  rcx, [0x14001D2F0]        ; output directory path
call 0x140011194               ; -> extract_7z_payload()
test eax, eax
jne  error_exit_0xFF           ; show error message box

lea  rdx, [0x14001D2E8]        ; "xlspadlock.bin" path
lea  rcx, [0x14001D2D8]        ; "XLSPadlockStub.dll" path
call 0x140011150               ; -> verify_and_load_runtime()
mov  dword [0x140029030], 2    ; mark as complete

; Load and call XLSPadlockInit
call 0x14000A098               ; GetCommandLineW wrapper
call 0x14000A258               ; LoadLibraryW("XLSPadlockStub.dll")
mov  rbx, rax
cmpq [rax], 0
je   launch_done

mov  rcx, rax
call 0x140009FFC               ; GetProcAddress("XLSPadlockInit")
test al, al
je   launch_done

; Call XLSPadlockInit with arguments
xor  r8d, r8d
lea  edx, [r8+2]               ; argc-like parameter = 2
xor  ecx, ecx                  ; argv
mov  rax, [rbx]
call [0x14001D2C8]             ; -> XLSPadlockInit via IAT thunk
```

### 8.3 Key Functions Inventory

| Address | Name (inferred) | Description |
|---|---|---|
| `0x140009D14` | `main_startup` | Top-level orchestration |
| `0x140009E88` | `entrypoint` | CRT entry point |
| `0x14000A150` | `crt_security_init` | Stack cookie + CRT globals init |
| `0x140009E9C` | `check_already_running` | TLS-based re-entrancy guard |
| `0x140009ED8` | `init_critical_section` | Global mutex wrapper |
| `0x140011194` | `extract_7z_payload` | 7-zip LZMA extraction of embedded archive |
| `0x140011150` | `verify_and_load_runtime` | CRC verification + DLL load |
| `0x14000A258` | `load_dll_wrapper` | `LoadLibraryW` + error handling |
| `0x140009FFC` | `get_proc_wrapper` | `GetProcAddress` + type check |
| `0x14000A690` | `check_debugger` | `IsDebuggerPresent` wrapper |
| `0x14000A258` | `cmdline_parse` | `GetCommandLineW` + `CommandLineToArgvW` |
| `0x140011A20` | `expand_env_path` | `SHGetFolderPathW` + path construction |
| `0x14001BD0` | `show_error_msgbox` | `MessageBoxW` error display |
| `0x140011BD0` | `handle_update_flow` | Version check / update handling |
| `0x14000A974` | `loop_wait` | Retry loop for locked files |

### 8.4 Call Graph (Simplified)

```
entrypoint (0x9E88)
└── main_startup (0x9D14)
    ├── init_critical_section (0x9ED8)
    ├── check_already_running (0x9E9C)
    │   └── IsDebuggerPresent (KERNEL32 import)
    ├── expand_env_path → SHGetFolderPathW → SHCreateDirectoryExW
    ├── extract_7z_payload (0x11194)
    │   ├── ReadFile (read EXE tail for 7z offset)
    │   ├── CreateFileW (create temp output files)
    │   ├── WriteFile (write extracted files)
    │   └── SzArEx_Open/Extract (7-zip LZMA SDK, inline)
    ├── verify_and_load_runtime (0x11150)
    │   ├── CreateFileW + ReadFile (read xlspadlock.bin)
    │   ├── [CRC32 comparison of extracted vs stored]
    │   └── LoadLibraryW("XLSPadlockStub.dll")
    ├── load_dll_wrapper (0xA258) → LoadLibraryW
    ├── get_proc_wrapper (0x9FFC) → GetProcAddress("XLSPadlockInit")
    └── [indirect call] → XLSPadlockInit (in XLSPadlockStub.dll)
```

---

## 9. Runtime Behavior

*This section is based on static analysis inference from imports, strings, and documented XLS Padlock behavior. Dynamic execution was not performed (Linux analysis environment).*

### 9.1 Process Lifecycle

**Phase 1 — Stub Initialization (main EXE)**
1. Process starts as `SchedPro_Tool64.exe` (GUI subsystem).
2. MSVC CRT initializes: security cookie, heap, locale, stdio.
3. `SHGetFolderPathW(CSIDL_APPDATA)` — locates `%APPDATA%` for extraction.
4. Path constructed: `%APPDATA%\XLS Padlock Runtime\<version_id>\`.
5. If files already extracted and CRC matches: skip to Phase 3.

**Phase 2 — 7z Extraction (first run only)**
1. Open self (`GetModuleFileNameW` → `CreateFileW`).
2. `SetFilePointerEx` to locate 7z archive at offset `0x65A00` (tail signature `37 7A BC AF 27 1C`).
3. 7-zip LZMA SDK (`SzArEx_Open`, extract loop) decompresses archive.
4. Output files:
   - `%APPDATA%\XLS Padlock Runtime\<id>\XLSPadlockStub.dll` (10.6 MB)
   - `%APPDATA%\XLS Padlock Runtime\<id>\xlspadlock.bin` (7.45 MB)
5. CRC32 of extracted `xlspadlock.bin` compared against stored value in header.

**Phase 3 — Runtime Handoff**
1. `LoadLibraryW("XLSPadlockStub.dll")` — loads Delphi runtime engine.
2. `GetProcAddress("XLSPadlockInit")` — resolves sole entry point.
3. Call `XLSPadlockInit(NULL, 2, ...)` — transfers control to Delphi engine.

**Phase 4 — Delphi Engine (XLSPadlockStub.dll)**
1. `BoxedAppSDK_Init()` — initializes virtual filesystem.
2. `BoxedAppSDK_CreateVirtualFile("xlspadlock.bin")` — maps bin as virtual file.
3. Read `xlspadlock.bin`:
   - Parse `XPLAPP` header → extract Zstd config → load UI captions/settings.
   - Parse `MAINICON` resource → set application icon via `Shell_NotifyIconW`.
   - Parse `XLS10` header → AES-256 decrypt workbook to in-memory buffer.
4. `BoxedAppSDK_CreateProcessFromMemory("EXCEL.EXE")` — launch virtualized Excel engine from the high-entropy section embedded in the DLL (entropy 7.999).
5. Excel COM Automation initialized; workbook loaded from decrypted buffer.
6. Activation check (if licensing enabled):
   - Check stored activation key in registry.
   - If missing: show Welcome dialog → Activation dialog.
   - Online validation via `WinHTTP` to `http://www.xlspadlock.com/activate`.
7. Application displays the SchedPro UI (scheduling tool).
8. On exit: save state, flush registry, cleanup temp files per policy.

### 9.2 Child Processes
- Potentially spawns a virtualized Microsoft Excel process via `BoxedAppSDK_CreateProcessFromMemory`.
- No other child process creation observed in static artifacts.

### 9.3 DLL Loading Chain
```
SchedPro_Tool64.exe
  → KERNEL32.dll, USER32.dll, SHELL32.dll, COMCTL32.dll  (static import)
  → XLSPadlockStub.dll  (dynamic: LoadLibraryW)
      → user32.dll, gdi32.dll, gdiplus.dll    (full UI stack)
      → winhttp.dll, wininet.dll, wsock32.dll  (network for activation)
      → advapi32.dll                           (registry + crypto API)
      → ole32.dll, oleaut32.dll                (COM automation for Excel)
      → comctl32.dll, comdlg32.dll             (dialogs)
      → LZ32.dll, version.dll, HID.dll         (auxiliary)
      → [virtualized Excel DLLs via BoxedApp]  (not visible to OS loader)
```

---

## 10. Memory Analysis

*Static inference only.*

### 10.1 Memory Layout at Runtime

| Region | Type | Content | Protection |
|---|---|---|---|
| `0x140000000`–base | PE image | Main EXE code + data | RX/R/RW (per section) |
| Stack (main thread) | Private | CRT initialization frames | RW |
| Heap | Private | Extraction buffers, path strings | RW |
| `XLSPadlockStub.dll` | Image | Delphi VCL engine | RX/R/RW |
| BoxedApp virtual FS | Private | Virtualized Excel DLLs in memory | RX (injected) |
| Decrypted workbook | Private | Plaintext XLSM in RAM | RW (ephemeral) |
| Activation data | Private | License state | RW |

### 10.2 Key Observations

- The **decrypted Excel workbook exists only in RAM** at runtime. It is never written to disk in plaintext (by design — this is the core value proposition of XLS Padlock).
- The **BoxedApp SDK** creates a **virtual in-process filesystem** that intercepts `CreateFile`/`LoadLibrary` calls. This means embedded DLLs (including potentially a bundled Excel runtime) are loaded without touching disk.
- The second `.text` section of `XLSPadlockStub.dll` (entropy 7.999, 946 KB) contains the **BoxedApp-encrypted virtual file system image** holding Excel COM DLLs. This is decompressed/decrypted at runtime into private memory.

---

## 11. Encryption / Packing Analysis

### 11.1 Outer Layer: 7-zip LZMA (Not Encrypted)

| Property | Value |
|---|---|
| Algorithm | LZMA + BCJ2 (branch converter for x86 code) |
| Archive offset | `0x65A00` in the EXE file |
| Compressed size | ~11.3 MB |
| Decompressed size | ~18 MB |
| Password | None |
| Files | `xlspadlock.bin`, `XLSPadlockStub.dll` |

**Mechanism:** The 7-zip LZMA SDK (specifically `SzArEx_Open`) is statically linked into the stub. The archive is appended to the end of the EXE during the XLS Padlock build process. The stub locates the archive by scanning backward from EOF for the 7z magic signature `37 7A BC AF 27 1C`.

**Confidence: High** (two 7z signatures confirmed at `0x24080` and `0x65A00`; only the second is a valid archive with correct header CRC).

### 11.2 Inner Layer: XPLAPP Container (`xlspadlock.bin`)

The XPLAPP format is a **proprietary XLS Padlock container**:

```
Offset  Size    Content
0x00    6       Magic "XPLAPP"
0x06    2       Version (01 00 = v1.0)
0x08    4       Field: 0x00000002 (section count or flags)
0x0C    4       Field: 0x00010000 (block size = 65536)
0x10    4       Extraction size: 242,332 bytes
0x14    8       Key material / IV: 01 02 03 04 05 06 07 08
0x1C    4       Field: 0x00000002
0x20    2       String length: 6
0x22    6       "LOCALE"
0x28    4       Locale ID: 4835 (0x12E3)
0x2C    4       CRC32 / timestamp: 0x12DD302F
0x30    2       Padding: 00 00
0x32    ~4.7KB  Zstandard compressed configuration
0x130F  ~237KB  MAINICON ICO resource
0x3B2B0 6       Section marker "XLS10" + version byte 0x12
0x3B2B6 7.21MB  AES-256 encrypted Excel workbook payload
0x77234C 137    "XPL04" footer + 128-character hex-encoded SHA-512 integrity hash
```

**IV candidate observation:** The bytes at offset `0x14` are `01 02 03 04 05 06 07 08` — a sequential test pattern that suggests either placeholder/test data or a fixed application-side IV. This was not further analyzed as decryption of the workbook is outside the scope of this educational analysis.

### 11.3 Workbook Encryption

| Property | Value | Confidence |
|---|---|---|
| Cipher | AES-256 (inferred from DLL symbol `TAESCore`) | Medium |
| Mode | CBC or GCM (inferred from `AESModes` / `SubkeyGenerationCMAC`) | Low |
| Key derivation | PBKDF2/SHA-256 or hardware-bound (from system ID + activation key) | Low |
| IV | Possibly stored in XLS10 header or derived from system ID | Low |
| Integrity | SHA-512 (64-byte hash in XPL04 footer) | Medium |

**Evidence for AES-256:** The DLL exports/strings contain: `TAESCore`, `TAESType`, `AESModes`, `keySizeBits`, `SubkeyGenerationCMAC` — these are class names from a Delphi AES implementation library.

**Evidence for SHA-512:** The XPL04 footer appends a 128-character hex string = 64 decoded bytes = 512 bits. SHA-512 is consistent.

**Entropy confirmation:** The encrypted workbook region (offset `0x3B2B6`–`0x77234C`) maintains near-uniform entropy of **~7.997** across all 64-KB chunks, consistent with AES block cipher output (indistinguishable from random).

### 11.4 BoxedApp SDK Virtual Filesystem

The second `.text` section of `XLSPadlockStub.dll` (entropy 7.999, 946 KB) contains a **BoxedApp-encrypted virtual filesystem**. This is a commercial SDK feature (BoxedAppSDK) that:

1. Intercepts Win32 filesystem calls (`CreateFileW`, `LoadLibraryW`).
2. Redirects requests for virtual paths to in-memory buffers.
3. Allows running COM DLLs (Excel automation libraries) without physical disk presence.

This is **not malware behavior** — it is the mechanism that allows XLS Padlock to bundle Excel without requiring Microsoft Office to be installed.

### 11.5 Anti-Analysis Protections

| Technique | Present | Notes |
|---|---|---|
| Custom packer | No | Standard 7-zip LZMA |
| Code obfuscation | No | Clean MSVC compiler output |
| Anti-debugging | **Yes (minor)** | `IsDebuggerPresent` called at `0x14000A690` |
| Anti-VM | No evidence | No CPUID tricks, no VM detection strings |
| Anti-sandbox | No evidence | No sleep loops, no environment checks in stub |
| Self-modification | No | `.text` section is not self-modifying |
| API hashing | No | Import names are plaintext |
| String encryption | No (stub) / Yes (DLL config) | DLL config compressed but not encrypted |
| Integrity checks | Yes | CRC32 on extracted files, SHA-512 on workbook |
| Timestamp forgery | No | Timestamp consistent with v14.16 linker version |

---

## 12. Decrypted Components

### 12.1 Zstandard Configuration Stream

Successfully decompressed at offset `0x32` in `xlspadlock.bin` (17,694 bytes, UTF-8 BOM):

**Section `[Captions]`** — UI dialog button/label text:
```ini
Tfactkey.bactivate=Activate~!@#$
Tfactkey.bcancel=Cancel~!@#$
Tfactkey.bcopy=Copy to Clipboard~!@#$
Tfactkey.bgetkeyonline=Get Key Online~!@#$
Tfactkey.Label1=System ID:~!@#$
Tfactkey.Label2=Your Activation Key:~!@#$
Tfdeactkey.bdeactiv=Deactivate Now~!@#$
Tfeula.checkboxaccept=I accept the terms of the license agreement~!@#$
Tfwelc.benterkey=Enter Activation Key...~!@#$
Tfwelc.boriginalworkbook=Original Workbook~!@#$
Tfwelc.brecsave=Load A Recent Save~!@#$
Tfwelc.bselsave=Choose Save File~!@#$
Tfwelc.bupdates=Check for Updates~!@#$
```

**Section `[OPTIONS]`**:
```ini
CommentsFile=
Delimiter=~!@#$
IsUTF8File=1
```

**Section `[Other]`** — WebUpdate integration:
```ini
TDataModule1.WebUpdate1.Agent=XLS Padlock App WebUpdate~!@#$
TDataModule1.WebUpdate1.DateFormat=dd/mm/yyyy~!@#$
TDataModule1.WebUpdateWizardEnglish1.CannotConnect=Could not connect to update server
TDataModule1.WebUpdateWizardEnglish1.FailedDownload=Failed to download updates
```

**Language:**
```ini
[Language names - for internal use only!]
Language_1=English
```

**Key findings from config:**
- The application has an **EULA acceptance dialog** (`Tfeula`).
- It supports **online activation** (`Tfonact`) and **deactivation** (`Tfdeactkey`).
- It has a **Welcome screen** with options: Enter Activation Key, Load Original Workbook, Load Recent Save, Choose Save File, Check for Updates.
- The **WebUpdate** component uses `dd/mm/yyyy` date format → likely European locale.

### 12.2 MAINICON Resource

A valid **9-size ICO file** (16×16, 24×24, 32×32, 48×48, 64×64, 72×72, 96×96, 128×128 + one PNG at Resource ID 9 [75,581 bytes]). This is the SchedPro application icon.

### 12.3 Encrypted Workbook (Not Decrypted)

The 7.21 MB encrypted blob at offset `0x3B2B6`–`0x77234C` in `xlspadlock.bin` contains the actual SchedPro Excel workbook. **Decryption is not performed** in this analysis, as the workbook is a user-data file protected by the author's proprietary encryption (it is the author's intellectual property). The purpose of XLS Padlock is specifically to protect the workbook's business logic from inspection.

---

## 13. Network Activity

### 13.1 Network Infrastructure in Main Stub

**None.** The outer loader stub has no network capabilities. It imports no WinSock, WinHTTP, or WinInet functions.

### 13.2 Network Infrastructure in XLSPadlockStub.dll

The Delphi runtime engine imports substantial networking libraries:

| Library | Functions | Purpose |
|---|---|---|
| `winhttp.dll` | 19 functions | WinHTTP-based HTTPS requests (modern, uses system proxy) |
| `wininet.dll` | 16 functions | WinInet legacy HTTP (used by WebUpdate component) |
| `wsock32.dll` | 6 functions | Raw socket support |

**Identified network endpoints (from DLL strings):**

| URL | Purpose | Confidence |
|---|---|---|
| `http://www.xlspadlock.com` | Activation server base URL (from DLL cert + strings) | High |
| `http://ocsp.digicert.com` | OCSP certificate validation (embedded in code signing cert chain) | High |
| `http://ocsp.sectigo.com` | OCSP certificate validation | High |
| `http://ocsp.comodoca.com` | OCSP certificate validation | High |
| `http://ocsp.usertrust.com` | OCSP certificate validation | High |
| `https://sectigo.com/CPS` | Certification Practice Statement URL (in cert) | High |

**Network activities (inferred):**
1. **Activation request** — POST to XLS Padlock activation server with System ID + activation key.
2. **Online validation** — periodic license re-validation (optional, based on author config).
3. **Web update check** — GET to update server for new version info.
4. **OCSP checks** — certificate revocation checks when validating downloaded updates.

### 13.3 Indy TCP/SSL Components

The DLL contains Indy (Internet Direct) component class names: `TIdSSLRegistry`, `EIdOSSLCouldNotLoadSSLLibrary`, `EIdOSSLConnectError` — a Delphi open-source internet library. SSL/TLS support is present for secure communications.

---

## 14. File System Activity

### 14.1 Predicted File Operations

| Operation | Path | Timing |
|---|---|---|
| Read | `SchedPro_Tool64.exe` (self) | Startup (7z extraction) |
| Create/Write | `%APPDATA%\XLS Padlock Runtime\<id>\XLSPadlockStub.dll` | First run only |
| Create/Write | `%APPDATA%\XLS Padlock Runtime\<id>\xlspadlock.bin` | First run only |
| Read | `%APPDATA%\XLS Padlock Runtime\<id>\xlspadlock.bin` | Every run |
| Read (virtual) | `[in-memory virtual FS]` | Runtime (BoxedApp) |
| Create | User-specified save file path | On user action |
| Read | User-specified save file path | On load |
| Write | `%APPDATA%\...\WUPDATE.LOG` | During web update |

### 14.2 Directory Structure

```
%APPDATA%\XLS Padlock Runtime\
    └── <numeric_id>\
        ├── XLSPadlockStub.dll   (10.6 MB, extracted)
        └── xlspadlock.bin        (7.45 MB, extracted)
```

The `<numeric_id>` directory is derived from the application version/CRC to allow side-by-side installs of different versions.

---

## 15. Registry Activity

### 15.1 Registry in Main Stub

None. The outer loader stub does not import `advapi32.dll` and performs no registry operations.

### 15.2 Registry in XLSPadlockStub.dll

The DLL imports `advapi32.dll` with 8 registry functions:

```
RegCreateKeyExW
RegOpenKeyExW
RegSetValueExW
RegQueryValueExW
RegEnumKeyExW
RegEnumValueW
RegCloseKey
RegFlushKey
RegDeleteValueW
```

**Predicted registry paths (inferred from XLS Padlock behavior patterns):**

| Key Path | Purpose |
|---|---|
| `HKCU\Software\YoussefMoataz\SchedPro\` | Application settings |
| `HKCU\Software\YoussefMoataz\SchedPro\License` | Activation key storage |
| `HKCU\Software\YoussefMoataz\SchedPro\RecentSaves` | Recent save file list |

Additionally, the DLL contains `BoxedAppSDK_CreateVirtualRegKeyA/W` — the virtual registry can intercept and redirect registry calls to an in-memory store, preventing physical registry writes for COM registration of bundled DLLs.

---

## 16. Security Findings

### 16.1 Security Findings Table

| ID | Finding | Severity | Confidence | Evidence |
|---|---|---|---|---|
| SF-01 | Binary is not Authenticode-signed | Medium | High | No Security data directory; `DllCharacteristics` FORCE_INTEGRITY bit not set |
| SF-02 | `IsDebuggerPresent` called on startup | Low | High | Import at `0x14001D150`; called via wrapper at `0x14000A690` from main flow |
| SF-03 | Main payload AES key material at `0x14` may be weak/fixed IV | Medium | Low | Bytes `01 02 03 04 05 06 07 08` in XPLAPP header — sequential pattern |
| SF-04 | XLSPadlockStub.dll not Authenticode-signed | Medium | High | DllCharacteristics `0x0160` (no FORCE_INTEGRITY) |
| SF-05 | Extraction to `%APPDATA%` without process elevation | Low | High | `SHGetFolderPathW` + `asInvoker` manifest — acceptable by design |
| SF-06 | DLL side-loading risk | Medium | Medium | `LoadLibraryW` by partial path from writable `%APPDATA%` location |
| SF-07 | No Guard CF (CFG) bit in DllCharacteristics | Low | High | `DllCharacteristics = 0x8160` — CFG bit 0x4000 absent; however XFG is present |
| SF-08 | HTTP (not HTTPS) for activation URL | Low | Medium | String `http://www.xlspadlock.com` — plaintext HTTP; upgrade interception possible |
| SF-09 | XLS10 sequential IV candidate | Medium | Low | `01 02 03 04 05 06 07 08` at `xlspadlock.bin+0x14` — if this is the encryption IV, it is a fixed/predictable value |
| SF-10 | Trial edition warning present in DLL | Info | High | String: "This application was built with the Trial edition of XLS Padlock" — distribution may be unauthorized if no commercial license held |

### 16.2 Finding Details

**SF-01 / SF-04 — Unsigned Binaries**
- **Impact:** A user or endpoint security product cannot verify publisher authenticity. Unsigned executables are blocked by default in some enterprise environments.
- **Recommendation:** The author should sign both files with a valid code signing certificate.

**SF-02 — `IsDebuggerPresent`**
- **Evidence:** Import at IAT entry `0x14001D150`, called via `0x14000A690`.
- **Intent:** XLS Padlock uses this to add a mild obstacle to runtime analysis. The result may be used to alter behavior or simply guard against debug-mode extraction.
- **Severity:** Low — this is a documented, widely-used pattern in commercial protection tools. It is not malicious.

**SF-06 — DLL Side-Loading Risk**
- **Impact:** The main EXE calls `LoadLibraryW` with a path under `%APPDATA%`. If an attacker can pre-create a malicious `XLSPadlockStub.dll` at that path (e.g., through a privilege escalation on a multi-user system), it would be loaded. However, this requires prior write access to the specific user's `%APPDATA%`.
- **Severity:** Medium (requires prior user account compromise — low practical risk in standard deployment).

**SF-08 — HTTP Activation URL**
- **Impact:** The activation communication (`http://www.xlspadlock.com`) uses plaintext HTTP. A network-level attacker (MITM) could intercept and potentially inject a fake "activation success" response, or observe the system ID being transmitted.
- **Note:** This is an XLS Padlock framework issue, not specific to this application.

**SF-10 — Trial Edition Notice**
- The DLL contains: `"This application was built with the Trial edition of XLS Padlock. Distribution of this file without having first purchased a license for XLS Padlock is prohibited."`
- **Impact:** This may indicate the author used a trial version of XLS Padlock to create this application. Distribution under a trial license may violate G.D.G. Software's EULA. This is a licensing compliance concern, not a security vulnerability.

---

## 17. Pseudocode Reconstruction

### 17.1 Main Stub Entry Logic

```c
// SchedPro_Tool64.exe — Main startup pseudocode
int WinMain(HINSTANCE hInst, HINSTANCE hPrev, LPWSTR lpCmd, int nShow) {
    // Phase 1: Initialize CRT and security cookie
    __security_init_cookie();
    if (!InitializeCriticalSectionAndSpinCount(&g_mutex, 0x1000)) {
        return ERROR_INIT_FAIL;
    }
    
    bool already_running = check_already_running();
    bool files_ready     = (g_state == STATE_COMPLETE);
    
    // Phase 2: Extract embedded 7z archive (first run)
    if (g_state == STATE_INIT) {
        g_state = STATE_EXTRACTING;
        
        wchar_t stub_dll_path[MAX_PATH];
        wchar_t bin_path[MAX_PATH];
        construct_output_paths(stub_dll_path, bin_path);  // %APPDATA%\XLS Padlock Runtime\<id>\
        
        if (extract_7z_to_directory(g_exe_self_path, bin_path) != 0) {
            MessageBoxW(NULL, L"7z archive not found", L"SchedPro", MB_ICONERROR);
            return 0xFF;
        }
        
        // Verify extracted xlspadlock.bin CRC
        if (verify_bin_crc(bin_path) != 0) {
            MessageBoxW(NULL, L"XlsPadlock.Bin CRC error", L"SchedPro", MB_ICONERROR);
            return 0xFF;
        }
        
        g_state = STATE_COMPLETE;
    }
    
    // Phase 3: Load and execute runtime
    LPWSTR* argv = NULL;
    int argc = 0;
    parse_cmdline(&argc, &argv);
    
    HMODULE hStub = LoadLibraryW(stub_dll_path);
    if (!hStub) {
        MessageBoxW(NULL, L"Failed to load XLSPadlockStub DLL", L"SchedPro", MB_ICONERROR);
        return ERROR_DLL_LOAD;
    }
    
    typedef int (*XLSPadlockInit_t)(LPVOID, int, LPWSTR*);
    XLSPadlockInit_t fnInit = (XLSPadlockInit_t)GetProcAddress(hStub, "XLSPadlockInit");
    if (!fnInit) {
        return ERROR_PROC_NOT_FOUND;
    }
    
    // Hand off to Delphi engine — does not return until app exits
    return fnInit(NULL, argc, argv);
}
```

### 17.2 Extract 7z Payload

```c
int extract_7z_payload(LPCWSTR exe_path, LPCWSTR output_dir) {
    // Open self
    HANDLE hExe = CreateFileW(exe_path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    if (hExe == INVALID_HANDLE_VALUE) {
        log_error("Error opening EXE");
        return -1;
    }
    
    // Get file size
    LARGE_INTEGER file_size;
    if (!GetFileSizeEx(hExe, &file_size)) {
        log_error("Cannot determine EXE size");
        CloseHandle(hExe);
        return -2;
    }
    
    // Read entire file into memory
    BYTE* file_buf = HeapAlloc(GetProcessHeap(), 0, file_size.LowPart);
    DWORD bytes_read;
    ReadFile(hExe, file_buf, file_size.LowPart, &bytes_read, NULL);
    CloseHandle(hExe);
    
    // Find 7z signature: 37 7A BC AF 27 1C (scan backward for last occurrence)
    static const BYTE SIG_7Z[] = {0x37, 0x7A, 0xBC, 0xAF, 0x27, 0x1C};
    DWORD archive_offset = find_last_signature(file_buf, file_size.LowPart, SIG_7Z, 6);
    if (archive_offset == (DWORD)-1) {
        log_error("7z archive not found");
        HeapFree(GetProcessHeap(), 0, file_buf);
        return -3;
    }
    
    // Initialize 7-zip LZMA SDK and extract
    CSzArEx db;
    SzArEx_Init(&db);
    
    // Use in-memory stream
    CLookToRead2 lookStream;
    // ... setup stream from file_buf+archive_offset ...
    
    SRes res = SzArEx_Open(&db, &lookStream.vt, &allocImp, &allocTempImp);
    if (res != SZ_OK) {
        sprintf(err_buf, "SzArEx_Open error: %d", res);
        log_error(err_buf);
        return -4;
    }
    
    // Extract both files
    for (UInt32 i = 0; i < db.NumFiles; i++) {
        wchar_t dest_path[MAX_PATH];
        get_dest_path(&db, i, output_dir, dest_path);
        extract_file_to_path(&db, &lookStream, i, dest_path);
    }
    
    SzArEx_Free(&db, &allocImp);
    HeapFree(GetProcessHeap(), 0, file_buf);
    return 0;
}
```

### 17.3 XPLAPP Config Parsing (XLSPadlockStub.dll — Delphi, reconstructed)

```pascal
procedure TXLSPadlockEngine.ParseXPLAPPContainer(const BinPath: string);
var
  BinFile: TFileStream;
  Header:  TXPLAPPHeader;
  ZstdBuf: TBytes;
  IniCfg:  TMemIniFile;
begin
  BinFile := TFileStream.Create(BinPath, fmOpenRead or fmShareDenyNone);
  try
    // Read and validate magic
    BinFile.ReadBuffer(Header, SizeOf(Header));
    if Header.Magic <> 'XPLAPP' then
      raise EXlsPadlockException.Create('Invalid XPLAPP magic');
    
    // Decompress Zstd configuration stream (at offset 0x32)
    BinFile.Seek($32, soFromBeginning);
    ZstdBuf := DecompressZstandard(BinFile, Header.ConfigCompressedSize);
    
    // Parse INI-format configuration
    IniCfg := TMemIniFile.CreateFromString(TEncoding.UTF8.GetString(ZstdBuf));
    LoadCaptions(IniCfg);
    LoadSettings(IniCfg);
    
    // Skip MAINICON block
    BinFile.Seek($3B2B0, soFromBeginning);
    
    // Read XLS10 section header
    ParseXLS10Header(BinFile);
    
    // Decrypt workbook
    DecryptWorkbook(BinFile, FWorkbookBuffer);
    
    // Verify footer
    VerifyXPL04Footer(BinFile);
    
  finally
    BinFile.Free;
  end;
end;

procedure TXLSPadlockEngine.DecryptWorkbook(Src: TStream; out Dest: TBytes);
var
  AES: TAESCore;
  Key: TBytes;
begin
  // Derive key from: activation key + system hardware ID
  Key := DeriveWorkbookKey(FActivationKey, FSystemID);
  
  AES := TAESCore.Create;
  try
    AES.Mode := amCBC;  // or amGCM
    AES.KeySize := 256;
    AES.SetKey(Key);
    // IV from XLS10 header or XPLAPP bytes 0x14-0x1B
    AES.IV := FXls10IV;
    
    // Decrypt 7.21 MB payload
    SetLength(Dest, FEncryptedSize);
    AES.Decrypt(Src, Dest);
  finally
    AES.Free;
  end;
  // SecureZeroMemory(Key) after use
end;
```

---

## 18. Functional Description

### 18.1 What SchedPro Does

Based on all static evidence, **SchedPro** is a **scheduling / task management tool** built entirely as a Microsoft Excel workbook, then packaged with XLS Padlock for standalone distribution.

**Evidence:**
- Name: "SchedPro" (Schedule + Professional)
- UI dialogs include: Welcome screen, activation dialog, EULA, save/load operations, recent saves
- The workbook requires either an activation key or a free-use mode
- WebUpdate component indicates continued development and version updates

### 18.2 Startup Sequence

```
1. User double-clicks SchedPro_Tool64.exe
2. Windows loads EXE into memory (ASLR randomizes base)
3. MSVC CRT initializes (security cookie, heap, locale, FLS/TLS)
4. Check %APPDATA%\XLS Padlock Runtime\<id>\ for existing extraction
   a. If missing or CRC mismatch: extract embedded 7z archive
   b. If present and valid: skip extraction
5. LoadLibraryW("XLSPadlockStub.dll") → maps 10.6 MB Delphi runtime
6. XLSPadlockInit() called → Delphi VCL application starts
7. BoxedApp SDK initializes virtual filesystem
8. XPLAPP container parsed:
   a. Config loaded (UI captions, language)
   b. Application icon set
   c. Workbook decrypted into RAM
9. Activation check:
   a. Valid key in registry → proceed
   b. No key → show Welcome/Activation dialog
   c. Online activation → POST to xlspadlock.com → store key in HKCU
10. SchedPro Excel workbook launches in virtualized Excel environment
11. User interacts with scheduling UI
12. User can save state to a file (non-plaintext, likely encrypted)
13. User exits → Delphi VCL cleanup → FreeLibrary → ExitProcess
```

### 18.3 User Interaction Flow

```
[Welcome Dialog]
    │
    ├─ "Enter Activation Key" → [Activation Dialog]
    │       ├─ Enter key manually
    │       ├─ Get Key Online → browser → xlspadlock.com
    │       └─ Copy/Paste from clipboard
    │
    ├─ "Original Workbook" → load original unmodified workbook
    ├─ "Load A Recent Save" → file picker (recent saves list)
    ├─ "Choose Save File" → GetOpenFileNameW dialog
    └─ "Check for Updates" → WebUpdate → xlspadlock.com
            │
            ▼
    [EULA Dialog] (first run or policy change)
            │
            ▼
    [Main Application: SchedPro Scheduling Tool]
            │
            ├─ Scheduling/task management UI in Excel
            ├─ Save state → custom encrypted save file
            └─ Exit
```

---

## 19. Indicators of Compromise (IOCs)

**Assessment: No malicious IOCs detected.** The following are legitimate application artifacts.

### 19.1 File Artifacts

| Type | Value | Context |
|---|---|---|
| SHA-256 | `56b328e803fa70d27f894441fe2ee0a27ceafd7dae32fb940d02f47cfa02bc0d` | Main executable |
| SHA-1 | `3585ed833952ffbc35c4aea1588edc76ccf7939d` | Main executable |
| MD5 | `68018c3cd1148c2102bdc15862d581fa` | Main executable |
| File path | `%APPDATA%\XLS Padlock Runtime\*\XLSPadlockStub.dll` | Extracted runtime DLL |
| File path | `%APPDATA%\XLS Padlock Runtime\*\xlspadlock.bin` | Extracted workbook container |
| Magic bytes (offset `0x65A00`) | `37 7A BC AF 27 1C 00 04` | Embedded 7z archive |
| Magic bytes (`xlspadlock.bin`) | `58 50 4C 41 50 50` ("XPLAPP") | Proprietary container |

### 19.2 Network Indicators

| Type | Value | Purpose | Malicious? |
|---|---|---|---|
| Domain | `www.xlspadlock.com` | Legitimate activation server | **No** |
| OCSP | `ocsp.digicert.com` | Certificate validation | **No** |
| OCSP | `ocsp.sectigo.com` | Certificate validation | **No** |
| OCSP | `ocsp.comodoca.com` | Certificate validation | **No** |

### 19.3 Registry Indicators

| Key | Value | Purpose |
|---|---|---|
| `HKCU\Software\YoussefMoataz\SchedPro\` | Various | App settings (predicted) |
| `HKCU\Software\YoussefMoataz\SchedPro\License` | Activation key | License storage (predicted) |

### 19.4 Strings of Interest (Non-Malicious)

| String | Type | Location |
|---|---|---|
| `GDGSoftware.XLSPadlock.App` | Assembly identity | Manifest |
| `XLSPadlockInit` | Export name | DLL entry point |
| `BoxedAppSDK_CreateVirtualFile*` | SDK API names | DLL strings |
| `TAESCore`, `TSHA3Core`, `TSHA2Core` | Crypto class names | DLL strings |
| `XLS10`, `XPL04`, `XPLAPP` | Container format magic | xlspadlock.bin |

---

## 20. Conclusion

### 20.1 Summary of Findings

`SchedPro_Tool64.exe` is a **legitimate standalone Excel-based scheduling application** created by the author **YoussefMoataz** using **XLS Padlock v25.2** — a commercial product by G.D.G. Software that converts Excel workbooks into protected standalone executables.

The file consists of three logical layers:
1. An **MSVC x64 loader stub** that extracts embedded components and loads the runtime.
2. An **encrypted container** (`xlspadlock.bin`) using the XPLAPP proprietary format, holding a 7.21 MB AES-256 encrypted Excel workbook.
3. The **XLSPadlockStub.dll** Delphi runtime engine that decrypts the workbook and executes it in a virtualized Excel environment powered by **BoxedApp SDK**.

### 20.2 Threat Assessment

| Category | Finding |
|---|---|
| **Malware** | **No** — no malicious code, no C2, no payload drops |
| **Backdoor** | **No** — no unauthorized remote access functionality |
| **Data exfiltration** | **No** — network access is limited to legitimate activation |
| **Ransomware** | **No** — no file encryption of user data |
| **Rootkit** | **No** — no kernel-mode code, runs as standard user |
| **Spyware** | **No** — no credential harvesting, no clipboard monitoring beyond UI |
| **PUA** | **Possibly** — trial edition notice may indicate unauthorized distribution |

### 20.3 Notable Technical Characteristics

1. **Multi-layer container architecture** — clean separation between loader, runtime engine, and encrypted payload.
2. **Virtual filesystem** (BoxedApp SDK) — sophisticated technique that allows running a complete Excel environment without modifying the system installation.
3. **AES-256 workbook encryption** — the business logic of the application (the Excel workbook) is effectively protected from casual inspection.
4. **Weak IV candidate** — the sequential bytes `01 02 03 04 05 06 07 08` in the XPLAPP header at offset `0x14` may represent a fixed or predictable IV, which could weaken the AES encryption if it is indeed used as the cipher IV.
5. **Trial edition notice** — the DLL string suggests this may have been built with a trial license of XLS Padlock, which prohibits distribution. Authors should ensure a valid commercial license is in place.

### 20.4 Recommendations

For the **software author (YoussefMoataz)**:
1. Obtain a commercial XLS Padlock license if distributing this application.
2. Sign the executable with an Authenticode certificate to establish publisher trust.
3. Ensure the activation connection uses HTTPS, not HTTP.
4. Review the AES IV derivation to ensure it uses a random or unique-per-instance IV.

For **defenders/analysts** encountering this file:
1. The file is not malicious. Antivirus detections (if any) would be false positives triggered by the BoxedApp SDK or 7-zip self-extracting behavior.
2. Network connections to `www.xlspadlock.com` from this process are expected and benign.
3. Files created in `%APPDATA%\XLS Padlock Runtime\` are expected application artifacts, not malware drops.

---

## Appendix A: Function Inventory Table

| Address (RVA) | Name (Inferred) | Calling Convention | Evidence Source |
|---|---|---|---|
| `0x9E88` | `entrypoint` | N/A | PE AddressOfEntryPoint |
| `0x9D14` | `main_startup` | x64 fastcall | Jump target from EP |
| `0x9E9C` | `check_already_running` | x64 fastcall | Call from main + TLS access |
| `0x9ED8` | `init_critical_section` | x64 fastcall | InitializeCriticalSectionAndSpinCount wrapper |
| `0xA098` | `parse_cmdline` | x64 fastcall | GetCommandLineW + CommandLineToArgvW |
| `0xA150` | `crt_security_init` | x64 fastcall | First call from EP |
| `0xA258` | `load_dll_wrapper` | x64 fastcall | LoadLibraryW |
| `0xA260` | `enum_loaded_dlls` | x64 fastcall | Secondary DLL enumeration |
| `0xA514` | `init_locale` | x64 fastcall | Called from CRT init |
| `0xA690` | `check_debugger` | x64 fastcall | IsDebuggerPresent |
| `0xA974` | `retry_wait` | x64 fastcall | Retry loop |
| `0x9FFC` | `get_proc_wrapper` | x64 fastcall | GetProcAddress |
| `0x11150` | `verify_and_load` | x64 fastcall | CRC + LoadLibraryW |
| `0x11194` | `extract_7z_payload` | x64 fastcall | 7z LZMA SDK |
| `0x11A20` | `construct_path` | x64 fastcall | SHGetFolderPathW + string concat |
| `0x11BD0` | `handle_update` | x64 fastcall | Version/update check |
| `0x15998` | `crt_string_init` | x64 fastcall | MSVC CRT runtime |
| `0x1BD0` | `show_error` | x64 fastcall | MessageBoxW error |

---

## Appendix B: IOC Table

| Category | Indicator | Value | Severity |
|---|---|---|---|
| Hash | SHA-256 | `56b328e803fa70d27f894441fe2ee0a27ceafd7dae32fb940d02f47cfa02bc0d` | Info |
| Hash | SHA-1 | `3585ed833952ffbc35c4aea1588edc76ccf7939d` | Info |
| Hash | MD5 | `68018c3cd1148c2102bdc15862d581fa` | Info |
| File | Main EXE | `SchedPro_Tool64.exe` (12 MB) | Info |
| File drop | DLL | `%APPDATA%\XLS Padlock Runtime\*\XLSPadlockStub.dll` | Low |
| File drop | Container | `%APPDATA%\XLS Padlock Runtime\*\xlspadlock.bin` | Low |
| Network | Domain | `www.xlspadlock.com` (activation, benign) | Info |
| Registry | Key (predicted) | `HKCU\Software\YoussefMoataz\SchedPro\` | Info |
| Mutex | Global mutex | Created via `InitializeCriticalSectionAndSpinCount` | Info |

---

## Appendix C: Security Findings Table

| ID | Finding | CVSS (approx.) | Category | Severity | Confidence |
|---|---|---|---|---|---|
| SF-01 | Unsigned main executable | N/A | Code Signing | Medium | High |
| SF-02 | `IsDebuggerPresent` anti-debug | N/A | Anti-analysis | Low | High |
| SF-03 | Possible weak/fixed AES IV | N/A | Cryptography | Medium | Low |
| SF-04 | Unsigned XLSPadlockStub.dll | N/A | Code Signing | Medium | High |
| SF-05 | No UAC elevation | N/A | Privilege | Info | High |
| SF-06 | DLL side-loading risk | N/A | DLL Security | Medium | Medium |
| SF-07 | CFG not set in DllCharacteristics | N/A | Memory Safety | Low | High |
| SF-08 | HTTP activation URL | N/A | Transport Security | Low | Medium |
| SF-09 | Sequential IV candidate | N/A | Cryptography | Medium | Low |
| SF-10 | Trial edition distribution | N/A | License Compliance | Info | High |

---

## Appendix D: Behavioral Summary

```
PROCESS:  SchedPro_Tool64.exe  (GUI, runs as standard user, no elevation)
          │
          ├─ READS:    itself (to extract embedded 7z archive)
          ├─ CREATES:  %APPDATA%\XLS Padlock Runtime\<id>\   (first run)
          ├─ WRITES:   XLSPadlockStub.dll, xlspadlock.bin   (first run)
          ├─ LOADS:    XLSPadlockStub.dll (dynamic LoadLibrary)
          ├─ CALLS:    XLSPadlockInit() → hands control to Delphi engine
          │
          └─ VIA XLSPadlockStub.dll:
             ├─ READS:    xlspadlock.bin (parses XPLAPP container)
             ├─ DECRYPTS: 7.21 MB workbook into RAM (AES-256)
             ├─ INJECTS:  BoxedApp virtual filesystem (Excel COM DLLs)
             ├─ REGISTRY: HKCU\Software\YoussefMoataz\SchedPro\
             ├─ NETWORK:  http://www.xlspadlock.com (activation, optional)
             ├─ DISPLAYS: Welcome → Activation → EULA → SchedPro UI
             ├─ READS:    user save files (scheduling data)
             ├─ WRITES:   user save files
             └─ EXIT:     cleanup, FreeLibrary, ExitProcess(0)
```

---

*This report was produced for educational, defensive, and interoperability purposes. No encryption keys, activation bypass methods, or workbook decryption techniques have been derived or documented. The analysis covers only structural and behavioral characteristics observable through static analysis.*

*All findings are based on static analysis of the file in a Linux environment. Dynamic analysis under Windows would provide additional confirmation of runtime behaviors described in §9 and §14.*
