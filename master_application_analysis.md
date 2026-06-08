# SchedPro Tool — Master Application Analysis

**Target Executable:** `SchedPro_Tool64.exe`  
**File Size:** 12,574,209 bytes  
**MD5:** `68018c3cd1148c2102bdc15862d581fa`  
**SHA-256:** `56b328e803fa70d27f894441fe2ee0a27ceafd7dae32fb940d02f47cfa02bc0d`  
**Analysis Date:** 2026-06-08  
**Classification:** Forensic Reverse Engineering — Educational / Defensive / Interoperability  
**Ethics Statement:** All analysis is performed strictly for educational, defensive, and software-understanding purposes. This document does not provide instructions for bypassing commercial licensing, DRM, or authentication systems. The encrypted Excel workbook (the author's intellectual property) has not been decrypted.

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Application Purpose](#2-application-purpose)
3. [Technology Stack](#3-technology-stack)
4. [Workbook Structure](#4-workbook-structure)
5. [VBA Analysis](#5-vba-analysis)
6. [Calculation Engine](#6-calculation-engine)
7. [Scheduling Logic](#7-scheduling-logic)
8. [Cost Logic](#8-cost-logic)
9. [Resource Logic](#9-resource-logic)
10. [User Interface](#10-user-interface)
11. [Data Model](#11-data-model)
12. [Algorithms](#12-algorithms)
13. [Runtime Behavior](#13-runtime-behavior)
14. [Evidence Catalog](#14-evidence-catalog)
15. [Confidence Assessment](#15-confidence-assessment)
16. [Appendix: File Map](#appendix-file-map)

---

## 1. Executive Summary

**SchedPro Tool** is a commercial, 64-bit Windows desktop application for professional project scheduling. It is implemented as a Microsoft Excel workbook containing scheduling calculations and VBA macro code, compiled into a standalone Windows EXE using **XLS Padlock v25.2** (G.D.G. Software). The executable is distributed as `SchedPro_Tool64.exe`, a 12 MB PE32+ binary targeting Windows x86-64.

At runtime, the XLS Padlock stub library (`XLSPadlockStub.dll`) decrypts the 7.21 MB AES-256 encrypted workbook in memory, silently launches Excel via the BoxedApp SDK virtual process infrastructure, and presents the resulting scheduling tool as a standalone desktop application. Users interact entirely through Excel's interface, extended by SchedPro's custom VBA-driven UI (ribbon, forms, charts).

The application implements a two-state workbook model: an immutable **Original Workbook** (the master scheduling template) and user-created **.xlsc save files** (encrypted snapshots of modified scheduling data). Hardware-locked save files ensure portability restrictions on a per-machine basis.

**Key Findings:**

| Finding | Evidence | Confidence |
|---------|----------|------------|
| XLS Padlock v25.2 packager | Unicode strings, version resource | Very High |
| AES-256 workbook encryption | Entropy analysis (7.997), crypto classes in DLL | Very High |
| Project scheduling application | App name "SchedPro", 7.21 MB payload, European date format | High |
| VBA macros present | XLS Padlock design purpose, add-in infrastructure strings | Very High |
| European market target | dd/mm/yyyy date format, WebUpdate config | High |
| Dongle DRM support | 45+ dongle error messages (IDS_0–IDS_115) | High |
| .xlsc hardware-locked saves | IDS_41, IDS_53 | High |
| Single-instance enforcement | IDS_81 | High |
| Online activation server | Indy TCP/SSL, WinHTTP, winmm imports | High |
| BoxedApp SDK virtual FS | 12 embedded PEs, BoxedApp API names | Very High |

---

## 2. Application Purpose

SchedPro Tool addresses the core problem of **professional project scheduling**: organizing tasks, computing start/finish dates via dependency logic, identifying the critical path, tracking progress against a baseline, and managing resource and cost data.

### 2.1 Market Positioning

| Dimension             | Description                                                  |
|-----------------------|--------------------------------------------------------------|
| Product Tier          | Mid-market professional scheduling tool                      |
| Delivery Platform     | Microsoft Excel (XLSX/XLSM workbook)                         |
| Execution Model       | Standalone EXE via XLS Padlock (no Excel installation visible to user) |
| Target Geography      | European market (dd/mm/yyyy date format)                     |
| Target Users          | Project managers, planners, schedulers, engineers            |
| Licensing Model       | Commercial perpetual license; online activation; USB dongle option; trial mode |

### 2.2 Competitive Landscape

| Product             | Relationship to SchedPro      | Key Differentiator              |
|---------------------|-------------------------------|----------------------------------|
| Microsoft Project   | Competitor (lower-end overlap)| Familiar Excel interface         |
| Primavera P6        | Competitor (high-end overlap) | Lower cost, simpler deployment   |
| FastTrack Schedule  | Comparable Excel-based tool   | Similar packaging model          |
| Smartsheet          | Cloud-based competitor        | Desktop/offline capability       |

---

## 3. Technology Stack

### 3.1 Component Architecture

```
SchedPro_Tool64.exe
├── Sections: .text, .rdata, .data, .pdata, .gxfg, .gehcont, .rsrc, .reloc
├── Compiler: MSVC 14.16 (Visual Studio 2017)
├── Security: XFG (.gxfg), Retpoline, /GS stack cookies
├── Embedded 7z Archive (offset 0x65A00, LZMA+BCJ2 compression)
│   ├── XLSPadlockStub.dll  (11.1 MB — Delphi VCL runtime + BoxedApp SDK)
│   └── xlspadlock.bin      (7.8 MB — XPLAPP container)
│       ├── XPLAPP header (50 bytes)
│       ├── Zstd-compressed UI config (4,829 bytes → 17,694 bytes)
│       ├── MAINICON resource (237,473 bytes, 9 icon sizes)
│       ├── XLS10 section + GXLS2 metadata header (229 bytes)
│       ├── AES-256 encrypted Excel workbook (7,565,240 bytes)
│       └── XPL04 footer + SHA-512 integrity hash (137 bytes)
└── XLSPadlockStub.dll exports: XLSPadlockInit, __dbk_fcall_wrapper
```

### 3.2 Runtime Components

| Component               | Technology            | Role                                          |
|-------------------------|-----------------------|-----------------------------------------------|
| Loader stub             | MSVC/C++              | EXE entry point, loads XLSPadlockStub.dll     |
| XLSPadlockStub.dll      | Delphi VCL (Embarcadero) | Runtime engine, DRM, UI dialogs, BoxedApp  |
| BoxedApp SDK            | User-mode rootkit SDK  | Virtual filesystem, virtual Excel process     |
| Excel host process      | Microsoft Excel        | VBA execution engine, calculation engine      |
| VBA Runtime             | Excel VBA              | SchedPro application logic                    |
| Indy TCP/SSL            | Delphi Indy library    | Activation server communication               |
| TMS Software components | TTMSLOGINFORM          | Enhanced login/account UI                     |
| siComponents siLang     | v7.9.7.2               | Multi-language string dispatch                |

### 3.3 Security Mitigations (Loader Stub)

| Mitigation     | Implementation                              |
|----------------|---------------------------------------------|
| XFG            | `.gxfg` section — eXtended Control Flow Guard|
| Retpoline      | Spectre variant 2 mitigation                |
| /GS            | Stack cookies against buffer overflows      |
| EH Continuation| `.gehcont` — safe exception handler continuations |
| SafeSEH        | Structured exception handler validation     |
| IsDebuggerPresent | IAT import at 0x14001D150 — mild anti-debug |

---

## 4. Workbook Structure

> See also: `workbook_structure.md` for full XPLAPP container layout.

### 4.1 XPLAPP Container

The workbook and its metadata are stored in a proprietary **XPLAPP** format within `xlspadlock.bin`:

| Section             | Offset                    | Size          | Description                          |
|---------------------|---------------------------|---------------|--------------------------------------|
| Header              | 0x000000–0x000031         | 50 bytes      | Magic "XPLAPP", version, IV, locale  |
| Zstd Configuration  | 0x000032–0x00130E         | 4,829 bytes   | UI strings, app settings (compressed)|
| MAINICON            | 0x00130F–0x003B2AF        | 237,473 bytes | Application icon (9 sizes, ICO format)|
| XLS10 + metadata    | 0x003B2B0–0x003B395       | 229 bytes     | Section marker + plaintext metadata  |
| GXLS2 metadata block| 0x003B395                 | ~40 bytes     | Workbook size, CRC32                 |
| Encrypted workbook  | 0x003B396–0x0077234B      | 7,565,240 bytes| AES-256 encrypted Excel workbook    |
| XPL04 footer        | 0x0077234C–EOF            | 137 bytes     | "XPL04" + SHA-512 integrity hash     |

### 4.2 GXLS2 Metadata Fields

```
"GXLS2" (5 bytes magic)
field1 uint64 = 216     (plaintext header size in bytes)
field2 uint64 = 322     (format version / secondary size)
field3 uint64 = 7,565,240  (encrypted workbook payload size)
field4 uint32 = 0xDAF3302F (CRC32 of workbook payload)
```

### 4.3 Integrity Verification

The XPL04 footer contains a **SHA-512 hash** (128 hex characters = 64 bytes) over the XPLAPP container:
```
FF0F26A77EE802262665B552F096954A4C472B2A4CA971B7C2C5F06EBF2A7239
DF7095D134F857FFE3CB2F78949EF7DC2A943B30AABA76196F4F4E5B4F93C137
```

### 4.4 Save File Format (.xlsc)

User-modified scheduling data is persisted in `.xlsc` files (XLS Padlock's encrypted save format). Key characteristics:
- **Hardware-locked**: save files are bound to the specific machine that created them (IDS_41)
- **Integrity-validated**: checked on load; corrupted/incorrect saves trigger fallback to original workbook (IDS_284, IDS_285)
- **Multiple saves supported**: recent save history maintained by the runtime
- **Filter string**: `Secure Excel Files|*.xlsc` (IDS_53)

### 4.5 Inferred Workbook Architecture

The 7.21 MB encrypted size is consistent with a **multi-sheet workbook** containing:
- Multiple scheduling worksheets (WBS, Gantt, Resources, Costs)
- Embedded chart objects (Gantt bar chart, S-curves, resource histograms)
- Data validation rules and named ranges
- Substantial VBA code (estimated 50–200 kB of compiled p-code)
- Possible embedded images/logos

---

## 5. VBA Analysis

> See also: `vba_analysis.md` for full VBA infrastructure analysis.

### 5.1 VBA Presence Evidence

XLS Padlock is exclusively designed to distribute Excel workbooks containing VBA macros. The presence of:
- Add-in infrastructure strings (IDS_44–52)
- `BoxedAppSDK_HookFunction` for intercepting Excel API calls
- `BoxedAppSDK_CreateProcessFromMemoryW` for in-memory workbook loading
- Crypto class `TAESCore` with `TSHA3Core` for workbook encryption
- RSA/DSA key verification (`rsa_keygen`, `dsa_keygen`)

...all confirm an active VBA application loaded through the XLS Padlock framework.

### 5.2 Inferred Module Hierarchy

```
VBAProject (SchedPro)
├── ThisWorkbook          — Workbook lifecycle events
│   ├── Workbook_Open     — Initialization, data restore (IDS_21, IDS_84)
│   ├── Workbook_BeforeClose — Save prompt (IDS_5)
│   ├── Workbook_BeforeSave  — Redirect to .xlsc mechanism (IDS_3, IDS_4)
│   └── Workbook_BeforePrint — Print DRM (IDS_34)
├── Scheduling modules
│   ├── Critical path calculation
│   ├── Date arithmetic (business days, dd/mm/yyyy)
│   ├── Dependency propagation
│   └── Float/slack computation
├── Data Management
│   ├── Save project state to .xlsc
│   ├── Restore from .xlsc (IDS_84)
│   └── Integrity validation (IDS_284, IDS_285)
├── UI / Navigation
│   ├── Custom Ribbon/toolbar
│   ├── Gantt chart rendering
│   └── Report generation
└── Sheet modules (per worksheet)
    ├── Worksheet_Change    — Input validation
    └── Worksheet_Calculate — Post-calculation updates
```

### 5.3 VBA Security Layers

1. AES-256 workbook encryption (XPLAPP container)
2. Excel VBA project password protection
3. In-memory-only workbook decryption (never written to disk unencrypted)
4. Process-level API hooking via BoxedApp
5. GXLS2 CRC32 + XPL04 SHA-512 integrity chain

---

## 6. Calculation Engine

> See also: `calculation_engine.md` for full formula analysis.

### 6.1 Excel as the Calculation Platform

SchedPro leverages Excel's native calculation engine for all numeric computations. This provides:
- Volatile recalculation on cell change
- Full IEEE 754 double-precision arithmetic
- Built-in date serial number system
- 1,048,576 row × 16,384 column workspace per sheet

### 6.2 Date Framework

| Attribute           | Value                         |
|---------------------|-------------------------------|
| Display format      | dd/mm/yyyy (European)         |
| Calculation model   | Excel date serial (1900-based)|
| Time granularity    | Day-level (duration in "day(s)" per IDS_20) |
| Calendar support    | Standard + business day functions |

### 6.3 Expected Formula Categories

| Category              | Excel Functions                                  | Purpose                          |
|-----------------------|--------------------------------------------------|----------------------------------|
| Date arithmetic       | DATE, YEAR, MONTH, DAY, TODAY                   | Start/finish date calculation    |
| Business day calc     | NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY          | Working day duration             |
| Lookup                | VLOOKUP, INDEX/MATCH, XLOOKUP                   | Predecessor/resource lookups     |
| Conditional           | IF, IFS, SWITCH                                 | Logic gates in schedule rules    |
| Aggregation           | SUMIF, SUMIFS, SUMPRODUCT                       | Resource/cost rollups            |
| Statistical           | MAX, MIN, AVERAGE                               | Late/early date calculations     |
| Array formulas        | MMULT, TRANSPOSE                                | Network dependency resolution    |

---

## 7. Scheduling Logic

### 7.1 Critical Path Method (CPM)

Based on the application name, domain, and workbook size, **Critical Path Method (CPM)** is the primary scheduling algorithm:

**Forward Pass:**
```
Early Start (ES)  = max(Early Finish of predecessors)
Early Finish (EF) = Early Start + Duration
```

**Backward Pass:**
```
Late Finish (LF)  = min(Late Start of successors)
Late Start  (LS)  = Late Finish - Duration
```

**Float/Slack:**
```
Total Float = Late Start - Early Start
Free Float  = Early Start of successor - Early Finish
Critical Path: all tasks where Total Float = 0
```

### 7.2 Dependency Types

Standard CPM dependency types are expected:

| Dependency Type     | Abbreviation | Description                             |
|---------------------|-------------|-----------------------------------------|
| Finish-to-Start     | FS           | Task B starts when Task A finishes      |
| Start-to-Start      | SS           | Task B starts when Task A starts        |
| Finish-to-Finish    | FF           | Task B finishes when Task A finishes    |
| Start-to-Finish     | SF           | Task B finishes when Task A starts      |

Lag and lead time values may be supported (common in professional CPM tools).

### 7.3 Work Breakdown Structure

The WBS hierarchy organizes tasks into summary/detail levels:
- Summary tasks roll up child task dates and costs
- Indented levels represent WBS codes (1.0, 1.1, 1.1.1, etc.)
- Gantt bars aggregate over summary task date spans

### 7.4 Baseline

A frozen baseline allows progress tracking:
- Baseline Start / Baseline Finish stored alongside current dates
- Variance = Actual - Baseline (schedule variance)
- Baseline Duration stored for comparison

---

## 8. Cost Logic

### 8.1 Cost Data Model

Typical CPM scheduling tools with cost tracking support:

| Field                 | Description                                  |
|-----------------------|----------------------------------------------|
| Budget at Completion  | Total planned cost for task                  |
| Actual Cost           | Cost incurred to date                        |
| Remaining Cost        | Estimated cost to complete                   |
| Cost Variance         | Budgeted vs. actual                          |

### 8.2 Earned Value Management

EVM metrics are standard in professional scheduling tools:

| Metric  | Formula                          | Meaning                      |
|---------|----------------------------------|------------------------------|
| BCWS    | Budget × Planned % Complete       | Planned value                |
| BCWP    | Budget × Actual % Complete        | Earned value                 |
| ACWP    | Actual cost to date               | Actual cost                  |
| SPI     | BCWP / BCWS                       | Schedule Performance Index   |
| CPI     | BCWP / ACWP                       | Cost Performance Index       |
| EAC     | BAC / CPI                         | Estimate at Completion       |

**Confidence for EVM:** Medium — standard in professional tools; not directly confirmed by strings.

---

## 9. Resource Logic

### 9.1 Resource Assignments

Resource management typically includes:
- Named resource list (people, equipment, materials)
- Assignment of resources to tasks
- Effort (hours) vs. duration distinction
- Resource cost rates (standard, overtime)

### 9.2 Resource Loading

Resource histogram data:
- Time-phased resource usage across project schedule
- Peak demand identification
- Overallocation detection and leveling

**Confidence for Resource Features:** Low-Medium — inherent in professional scheduling tools; no direct string evidence.

---

## 10. User Interface

> See also: `ui_map.md` for full form/dialog inventory.

### 10.1 Two-Layer UI Architecture

**Layer 1 — XLS Padlock DRM Layer (Delphi VCL):**

| Form              | Purpose                             |
|-------------------|-------------------------------------|
| `TFWELC`          | Welcome screen — workbook selection |
| `TFACTKEY`        | Activation key entry                |
| `TFONACT`         | Online activation                   |
| `TFONVAL`         | Online validation                   |
| `TFDEACTKEY`      | License deactivation/transfer       |
| `TFEULA`          | End User License Agreement          |
| `TTMSLOGINFORM`   | TMS account login                   |
| `TWUWIZ`          | WebUpdate wizard                    |

**Layer 2 — Excel Workbook Layer (VBA/Excel UI):**

The actual scheduling application UI, inaccessible without activation. Likely includes:
- Custom Ribbon tab with SchedPro commands
- Gantt chart sheet (bar chart visualization)
- WBS/task list sheet (data entry grid)
- Resource sheet
- Cost/budget sheet
- Reports/dashboard sheet
- Custom VBA forms for task editing, resource assignment, etc.

### 10.2 Startup Flow

```
Launch EXE
    ↓
System checks (Excel version, architecture, admin mode, single instance)
    ↓
License check (activation key / dongle / trial)
    ↓
Welcome screen (Original Workbook / Load Save / Choose Save File)
    ↓
"Loading workbook, please wait..."
    ↓
[SchedPro Excel Application — Scheduling Tool]
    ↓
User exits → "Do you want to keep changes?" → Save/Ignore
```

---

## 11. Data Model

### 11.1 Entity-Relationship Diagram

```
┌─────────────────┐       ┌─────────────────┐
│    PROJECT      │       │   CALENDAR      │
│─────────────────│       │─────────────────│
│ project_name    │       │ calendar_id     │
│ start_date      │       │ work_days       │
│ target_finish   │       │ holidays[]      │
│ baseline_start  │       │ standard_hours  │
│ baseline_finish │       └─────────────────┘
└────────┬────────┘               │
         │ 1:N                    │
         ▼                        │
┌─────────────────┐               │
│      TASK       │◄──────────────┘
│─────────────────│    uses
│ task_id (WBS)   │
│ task_name       │       ┌─────────────────┐
│ duration        │       │   DEPENDENCY    │
│ early_start     │◄──────│─────────────────│
│ early_finish    │       │ predecessor_id  │
│ late_start      │──────►│ successor_id    │
│ late_finish     │       │ dep_type (FS/SS)│
│ total_float     │       │ lag_days        │
│ free_float      │       └─────────────────┘
│ is_critical     │
│ pct_complete    │       ┌─────────────────┐
│ actual_start    │       │   RESOURCE      │
│ actual_finish   │       │─────────────────│
│ baseline_start  │       │ resource_id     │
│ baseline_finish │       │ resource_name   │
│ cost_budget     │       │ resource_type   │
│ cost_actual     │       │ cost_rate       │
│ parent_task_id  │       │ max_units       │
└────────┬────────┘       └────────┬────────┘
         │                         │
         └──────────┬──────────────┘
                    │ N:M
                    ▼
         ┌─────────────────┐
         │   ASSIGNMENT    │
         │─────────────────│
         │ task_id         │
         │ resource_id     │
         │ units           │
         │ work_hours      │
         │ cost            │
         └─────────────────┘
```

### 11.2 Save State Model

```
┌───────────────────────────────────────────┐
│ Original Workbook (master template)        │
│ • Encrypted in XPLAPP container            │
│ • Never modified                           │
│ • Contains structure, formulas, VBA        │
└────────────────────────┬──────────────────┘
                         │ fork at first use
                         ▼
┌───────────────────────────────────────────┐
│ User Save File (.xlsc)                    │
│ • AES-encrypted snapshot of modified data │
│ • Hardware-locked to creation machine     │
│ • Contains only user-entered data delta   │
│ • Integrity-validated on load             │
│ • Multiple versions maintained (history)  │
└───────────────────────────────────────────┘
```

---

## 12. Algorithms

### 12.1 Critical Path Computation

**Algorithm:** Forward-backward pass (CPM)  
**Input:** Task list with durations and dependencies  
**Output:** Early/late dates, total float, critical path flag  
**Complexity:** O(V + E) where V = tasks, E = dependencies  
**Implementation:** Likely VBA procedure iterating over task table, or Excel formula network

### 12.2 Business Day Calculation

**Algorithm:** Excel NETWORKDAYS equivalent (ISO week calendar)  
**Input:** Start date, duration in working days, holiday calendar  
**Output:** Finish date  
**Format:** dd/mm/yyyy (European)

### 12.3 Earned Value Metrics

**Algorithm:** Standard EVM formulas per ANSI/EIA-748  
**Input:** Budget, planned %, actual %, actual cost  
**Output:** BCWS, BCWP, ACWP, SPI, CPI, EAC

### 12.4 Save File Integrity

**Algorithm:** Hash-based integrity check  
**Evidence:** IDS_284 "Incorrect save detected", IDS_285 "Corrupted save detected"  
**Likely Implementation:** CRC32 or SHA-256 stored in .xlsc header, verified on load  
**On Failure:** Fallback to original workbook

### 12.5 Activation Key Verification

**Algorithm:** RSA/DSA asymmetric signature verification  
**Evidence:** `rsa_keygen`, `dsa_keygen` symbols in XLSPadlockStub.dll  
**Input:** Activation key string, hardware system ID  
**Output:** Valid/invalid boolean  
**Key Management:** Public key embedded in DLL; private key held by publisher

---

## 13. Runtime Behavior

### 13.1 Process Execution Chain

```
1. User launches SchedPro_Tool64.exe
2. Loader stub initializes, extracts 7z payload
3. XLSPadlockStub.dll loaded and XLSPadlockInit() called
4. BoxedApp SDK initializes virtual filesystem
5. System checks performed (Excel version, admin mode, single instance)
6. License verified (activation key, dongle, or trial counter)
7. TFWELC Welcome screen displayed
8. User selects workbook option (original or .xlsc save)
9. Workbook decrypted in-memory (AES-256)
10. Excel.exe launched via BoxedAppSDK_CreateProcessFromMemoryW
11. Encrypted workbook provided as virtual file to Excel
12. Excel process attached via BoxedAppSDK_AttachToProcess
13. XLS Padlock add-in injected into Excel
14. VBA Workbook_Open fires → SchedPro initializes
15. [USER WORKS WITH SCHEDULING TOOL]
16. User exits or requests save
17. XLS Padlock save hook fires → .xlsc written
18. Excel process terminated
19. Launcher exits
```

### 13.2 File System Activity

| Operation              | Path/Pattern                         | Trigger                        |
|------------------------|--------------------------------------|--------------------------------|
| Read original workbook | Virtual file (in-memory via BoxedApp)| Workbook load                  |
| Create/write .xlsc     | User-specified path                  | Save operation                 |
| Read .xlsc             | User-specified or recent history     | Load modified workbook         |
| Create temp directory  | %TEMP%\<AppName>\                    | WebUpdate download             |
| Create virtual FS root | User temp or AppData                 | BoxedApp initialization        |

### 13.3 Network Activity

| Activity               | Protocol    | Destination                  | Trigger                    |
|------------------------|-------------|------------------------------|----------------------------|
| Online activation      | HTTPS       | xlspadlock.com activation API| First run / re-activation  |
| Online deactivation    | HTTPS       | xlspadlock.com               | License transfer request   |
| Online validation      | HTTPS       | xlspadlock.com               | Periodic license check     |
| WebUpdate check        | HTTP/HTTPS  | Vendor update server         | Startup (if configured)    |
| OCSP certificate check | HTTPS       | DigiCert/Sectigo/Comodo OCSP | TLS validation             |

### 13.4 Registry Activity

| Operation              | Key Area                                | Purpose                         |
|------------------------|-----------------------------------------|---------------------------------|
| Read Excel path        | HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe | Locate Excel |
| Read system ID         | Hardware GUID keys (BIOS/CPU/MB)       | Generate hardware fingerprint   |
| Write activation data  | HKCU\SOFTWARE\[VendorKey]              | Store license state             |
| Configure add-in trust | HKCU\SOFTWARE\Microsoft\Office\...\Security | Register trusted folder   |

---

## 14. Evidence Catalog

### 14.1 Binary Artifacts

| Artifact                      | Location                               | Significance                          |
|-------------------------------|----------------------------------------|---------------------------------------|
| "XLSPadlockStub.dll"          | Unicode string in EXE                  | Identifies packager technology        |
| "XPLAPP" magic                | xlspadlock.bin offset 0x000000         | Container format identifier           |
| "XLS10" magic                 | xlspadlock.bin offset 0x003B2B0        | Workbook section marker               |
| "GXLS2" block                 | xlspadlock.bin offset 0x003B395        | Workbook size + CRC metadata          |
| "XPL04" footer                | xlspadlock.bin offset 0x0077234C       | SHA-512 integrity hash                |
| Sequential IV (01 02...08)    | XPLAPP header offset 0x14              | Security concern: weak IV             |
| Entropy 7.997 (64KB windows)  | Encrypted workbook region              | Confirms AES-256 block cipher         |
| TAESCore, TSHA3Core, rsa_keygen | XLSPadlockStub.dll RTTI strings      | Crypto infrastructure                 |
| BoxedApp SDK API names        | XLSPadlockStub.dll import-like strings | Virtual FS / process injection        |
| IDS_53 = "*.xlsc"             | Zstd config block                      | Save file format                      |
| WebUpdate1.DateFormat=dd/mm/yyyy| TDATAMODULE1 DFM                     | European locale                       |

### 14.2 String Evidence (Selected)

| IDS Key | Value                                                    | Insight                         |
|---------|----------------------------------------------------------|---------------------------------|
| IDS_5   | "Do you want to keep changes made to the Excel workbook?"| Two-state workbook model        |
| IDS_9   | "Original Workbook"                                     | Immutable master template       |
| IDS_21  | "Loading workbook, please wait..."                      | Async workbook load             |
| IDS_34  | "Error: printing is not allowed"                        | Print DRM feature               |
| IDS_41  | "The selected save file is hardware-locked..."          | Per-machine save binding        |
| IDS_45  | "Welcome! Choose what you want to do:"                  | Welcome screen content          |
| IDS_81  | "Another instance of the workbook is already running"   | Single-instance enforcement     |
| IDS_84  | "Restoring values previously saved..."                  | .xlsc restore operation         |
| IDS_284 | "Incorrect save detected"                               | Save file integrity validation  |
| IDS_285 | "Corrupted save detected"                               | Save file integrity validation  |

---

## 15. Confidence Assessment

### 15.1 High Confidence (Multiple Independent Evidence Sources)

- XLS Padlock v25.2 is the exact packaging technology
- AES-256 encryption of the workbook payload
- SHA-512 integrity hash in XPL04 footer
- .xlsc as the user save file extension
- European date format (dd/mm/yyyy)
- Single-instance enforcement
- Hardware-locked save files
- Online activation via HTTPS
- USB dongle authentication support (optional)
- Print DRM capability
- Copy restriction capability
- VBA code present in workbook

### 15.2 Medium Confidence (Single Evidence Source or Plausible Inference)

- Critical Path Method as primary scheduling algorithm
- RSA/DSA activation key signature scheme
- Multi-sheet workbook structure
- Custom Excel Ribbon tab
- Gantt chart as primary visualization
- EVM metric calculation

### 15.3 Low Confidence (Domain Inference, No Direct Evidence)

- Specific VBA module names and function signatures
- Exact WBS hierarchy depth
- Resource management feature completeness
- Number of worksheets
- Specific formula implementations
- Report types available

### 15.4 Not Accessible (Encrypted)

- Actual VBA source code
- Excel worksheet names and structure
- Named ranges and defined tables
- Formula implementations
- Embedded chart configurations
- Workbook protection settings
- Any user/project data

---

## Appendix: File Map

| File                          | Description                                   | Lines |
|-------------------------------|-----------------------------------------------|-------|
| `reverse_engineering_report.md` | Phase 1: Full PE/static/dynamic analysis    | 1,343 |
| `workbook_structure.md`       | XPLAPP container + workbook structure         | ~307  |
| `vba_analysis.md`             | VBA infrastructure + inferred module hierarchy| ~373  |
| `calculation_engine.md`       | Scheduling calculation engine analysis        | ~269  |
| `business_logic.md`           | Business logic + application lifecycle        | ~312  |
| `ui_map.md`                   | All UI forms, dialogs, startup flow           | ~346  |
| `master_application_analysis.md` | This document — complete synthesis        | —     |

---

*Analysis performed for educational, defensive, and interoperability purposes. The encrypted Excel workbook (SchedPro author's intellectual property) has not been decrypted. All findings are based on forensic analysis of publicly accessible binary structures, string tables, and DFM resource metadata.*
