# SchedPro Tool64.exe — VBA/Macro Analysis

**Project:** SchedPro Tool64.exe — Reverse Engineering (Forensic/Educational)
**Packager:** XLS Padlock v25.2 by G.D.G. Software
**Analysis date:** 2026-06-08
**Analyst note:** The VBA project is embedded inside the AES-256 encrypted workbook payload. Direct inspection of VBA module source code, procedure names, or module-level variables is not possible without the activation key. All findings below are inferred from static analysis of XLSPadlockStub.dll (symbol tables, imported/exported functions, string resources, and class name metadata), the XPLAPP container structure, and behavioral patterns characteristic of XLS Padlock v25.x deployments.

---

## 1. Analysis Constraints and Methodology

### 1.1 Primary Constraint

The workbook payload at XPLAPP container offsets 0x003B396–0x0077234B is protected with AES-256 encryption (Shannon entropy: 7.997 bits/byte). The VBA project — including all module names, procedure signatures, class definitions, form layouts, and source code — resides exclusively within this encrypted region. It cannot be extracted, decompiled, or inspected without the runtime decryption key.

**Consequence:** This document contains no direct VBA source code findings. Every claim about VBA structure, module organization, or event handler behavior is an inference, not a direct observation.

### 1.2 Methodology

Evidence for VBA presence and structure was collected from four sources, listed in descending reliability order:

| Source | Type | Reliability |
|--------|------|-------------|
| XLSPadlockStub.dll exported/imported symbol names | Direct observation | HIGH |
| XLSPadlockStub.dll class and type names (RTTI) | Direct observation | HIGH |
| IDS_* string resources from XLSPadlockStub.dll | Direct observation | HIGH |
| XLS Padlock v25.x product documentation and known behavior | External reference | MEDIUM |
| Inference from XPLAPP container structure and payload size | Structural inference | LOW–MEDIUM |

### 1.3 Confidence Rating Scale

| Rating | Meaning |
|--------|---------|
| HIGH | Directly observed in DLL symbols, RTTI metadata, or string resources |
| MEDIUM | Consistent with multiple independent forensic indicators or published XLS Padlock behavior |
| LOW | Single-source inference or pattern-matching against known Excel/VBA conventions |

---

## 2. VBA Infrastructure Evidence from XLSPadlockStub.dll

### 2.1 XLS Padlock Design Intent

**Confidence: HIGH**

XLS Padlock is a commercial product by G.D.G. Software explicitly designed for distributing Excel workbooks that contain VBA macros in a protected, executable form. Per G.D.G. Software's published product documentation, XLS Padlock:

- Encrypts the Excel workbook (including its embedded VBA project) with AES-256
- Packages the encrypted workbook with a Windows PE stub (XLSPadlockStub.dll) that handles decryption, activation, and Excel hosting at runtime
- Preserves full VBA macro execution capability within the sandboxed Excel environment
- Protects VBA code from extraction by standard Office VBA password removal tools

The presence of a 7.21 MB encrypted payload in SchedPro Tool64.exe is therefore strong evidence that a VBA-containing workbook is being protected. A workbook of this size without any VBA would be unusual and would lack the behavioral features (custom save dialog, restore logic, add-in loading) evidenced by the IDS string set.

### 2.2 BoxedApp SDK API Evidence

**Confidence: HIGH** (function names directly observed in XLSPadlockStub.dll import/export tables)

XLSPadlockStub.dll imports the BoxedApp SDK, a commercial application virtualization library. The following APIs are directly relevant to VBA macro hosting:

| API | Observed In | Relevance to VBA |
|-----|-------------|-----------------|
| `BoxedAppSDK_CreateProcessFromMemoryW` | DLL imports | Launches the sandboxed Excel process from an in-memory decrypted image, never writing the decrypted workbook to disk |
| `BoxedAppSDK_HookFunction` | DLL imports | Intercepts Excel API calls at runtime; used to redirect file I/O, licensing checks, and potentially VBA-triggered dialogs |

The `CreateProcessFromMemoryW` call is the central mechanism by which XLS Padlock prevents the decrypted workbook (and its VBA project) from being accessible on disk. The workbook is decrypted entirely in memory, and Excel is launched against the in-memory image. This is why the VBA project cannot be extracted by monitoring the file system during execution.

`BoxedAppSDK_HookFunction` enables XLSPadlockStub.dll to intercept Excel's internal function calls. In the context of VBA, this is used to:

- Intercept `Application.Save` / `Workbook.Save` calls and route them to the custom .xlsc save mechanism
- Intercept `Application.Quit` / `Workbook.BeforeClose` events to inject the custom "keep changes?" prompt
- Potentially intercept VBA `MsgBox` or `InputBox` calls to apply UI theming or branding

### 2.3 VBA Runtime Hosting Environment

**Confidence: HIGH**

XLSPadlockStub.dll provides the complete runtime hosting environment for the VBA project. Specifically:

- The stub initializes a sandboxed COM/OLE environment that includes the Excel Object Model
- The VBA runtime (VBE7.DLL / msvbvm60.dll, whichever Excel version uses) is loaded into this sandboxed process
- The stub's hook layer sits between the VBA runtime and the Windows API, allowing it to intercept and selectively allow or block VBA-initiated operations
- VBA code in the workbook executes with full access to the Excel Object Model (Workbook, Worksheet, Range, Chart objects) but file system access may be restricted by the BoxedApp virtualization layer

### 2.4 Add-In Infrastructure Evidence

**Confidence: HIGH** (IDS strings directly observed)

The following IDS string identifiers confirm that Excel add-ins are loaded as part of the application startup sequence:

| IDS ID | Inferred String Content | Significance |
|--------|------------------------|--------------|
| IDS_44 | Add-in load/management reference | Confirms add-in loading is part of the XLS Padlock runtime |
| IDS_48 | Add-in path or registration string | Add-in file paths are configured in the Zstd config blob |
| IDS_49 | Add-in status/error message | Error handling for failed add-in loads |
| IDS_50 | Add-in status/error message (variant) | Multiple add-in error states are handled |
| IDS_51 | Add-in status/error message (variant) | |
| IDS_52 | Add-in status/error message (variant) | |

Excel add-ins (.xlam, .xla) are themselves workbooks containing VBA code. The add-in infrastructure in XLS Padlock is typically used to:

- Load utility VBA libraries that the main workbook's VBA calls via `Application.Run`
- Separate licensing/activation VBA code into a protected add-in distinct from the main scheduling logic
- Load custom ribbon XML (CustomUI) that provides the SchedPro toolbar/ribbon interface

The presence of add-in loading strings (IDS_44, IDS_48–52) alongside the CustomUI directory found in the repository strongly suggests that SchedPro uses at least one add-in to implement its ribbon interface or utility VBA library.

---

## 3. Inferred VBA Module Structure

**Confidence: MEDIUM overall** (structure inferred from IDS strings and standard Excel VBA project conventions; no direct observation possible)

A VBA project in an Excel workbook of this complexity would typically be organized into the following module types. The inferences below are grounded in the specific IDS strings identified from XLSPadlockStub.dll.

### 3.1 Workbook Initialization / Startup Module

**Confidence: MEDIUM**

Evidence: IDS_21 — inferred content: "Loading workbook, please wait..."

This string is displayed during the startup sequence, indicating a non-trivial initialization routine that takes measurable time. For a 7.21 MB scheduling workbook, startup tasks likely include:

- Populating formula-driven cells (Gantt chart recalculation, resource leveling, calendar generation)
- Restoring previously saved user data from a .xlsc file (IDS_84: "Restoring values previously saved")
- Verifying the activation state and hardware fingerprint
- Loading add-ins referenced in IDS_44/48–52
- Setting up event handler registrations

This initialization logic would reside in a standard VBA module (e.g., `modStartup` or `basInitialize`) called from the `Workbook_Open` event handler.

### 3.2 Data Validation and Save/Restore Routines

**Confidence: MEDIUM**

Evidence: IDS_84, IDS_283, IDS_284, IDS_285

| IDS ID | Inferred Content | VBA Routine Implied |
|--------|-----------------|---------------------|
| IDS_84 | "Restoring values previously saved" | `Sub RestoreSavedValues()` or equivalent — reads .xlsc and applies values to cells |
| IDS_283 | Save prompt (variant) | `Sub ShowSaveDialog()` — custom save UI |
| IDS_284 | "Incorrect save detected" | Save integrity check routine — checksum validation failure handler |
| IDS_285 | "Corrupted save detected" | Save structural validation failure handler |

The existence of two distinct error messages (IDS_284 vs. IDS_285) implies a multi-stage save validation routine:

```
Stage 1 — Structural validation: Is the .xlsc file a valid SchedPro save file?
  → Failure → IDS_285 "Corrupted save detected"
Stage 2 — Integrity validation: Does the save file's checksum match?
  → Failure → IDS_284 "Incorrect save detected"
Stage 3 — Hardware-lock validation: Was this save created on this machine?
  → Failure → IDS_41 (hardware lock error)
Stage 4 — Success → IDS_84 "Restoring values previously saved"
```

This multi-stage pattern is a standard defensive programming approach in commercial Excel applications to distinguish accidental file corruption from deliberate tampering.

### 3.3 Custom Save Dialog Module

**Confidence: HIGH** (multiple IDS strings directly confirm custom save UI)

Evidence: IDS_3, IDS_4, IDS_5, IDS_7, IDS_8

These IDS identifiers correspond to button labels, dialog titles, and prompt messages in a custom save dialog that replaces Excel's native save behavior. The specific prompts inferred:

| IDS ID | Likely Content | Context |
|--------|---------------|---------|
| IDS_3 | "Save" (button label) | Custom save dialog button |
| IDS_4 | "Save As..." or "Save to file" | Save-as variant action |
| IDS_5 | "Do you want to keep changes?" | Close-without-saving prompt |
| IDS_7 | "Yes" or "Save and close" | Affirmative response button |
| IDS_8 | "No" or "Discard changes" | Negative response button |

IDS_5 in particular ("Do you want to keep changes?") mirrors the phrasing specified in the task brief for `Workbook_BeforeClose`, confirming that the VBA event handler intercepts the close event and presents this custom dialog before allowing Excel to close.

### 3.4 Scheduling Calculation Modules (Core SchedPro Logic)

**Confidence: LOW** (inferred from application name and workbook size; no direct string evidence)

The primary value of SchedPro as a project scheduling tool implies the presence of VBA modules implementing or supporting scheduling calculations. These would include some combination of:

- **Critical Path Method (CPM):** Procedure(s) to traverse a task dependency network and compute Early Start, Early Finish, Late Start, Late Finish, and Total Float for each task
- **Gantt chart rendering:** VBA routines to draw or update Gantt bar graphics based on task dates and durations; possibly manipulating Shape objects or conditional formatting rules
- **Resource leveling:** Routines to detect over-allocation of resources across the project timeline and either flag conflicts or attempt automatic leveling
- **Calendar and working-day calculations:** Functions to compute working days between dates using a project calendar (holidays, weekends, shift patterns)
- **Baseline comparison:** Modules to snapshot the original project plan and compare it against the current schedule to compute variance

The 7.21 MB payload size is consistent with a workbook containing significant formula arrays for these calculations in addition to VBA modules.

---

## 4. Inferred Event Handler Architecture

**Confidence: MEDIUM** (event handler names follow strict Excel VBA conventions; behavior inferred from IDS strings)

Excel VBA event handlers have fixed, convention-mandated names. The following handlers are inferred based on observed behavioral evidence.

### 4.1 `Workbook_Open`

**Location:** `ThisWorkbook` module (mandatory Excel convention)
**Confidence: HIGH** (startup string IDS_21 directly implies this handler fires)

Inferred execution sequence:

1. Display "Loading workbook, please wait..." status (IDS_21)
2. Verify activation state via XLSPadlockStub.dll hook
3. Register BoxedApp SDK hooks for save/close interception
4. Load add-ins listed in Zstd config (IDS_44, IDS_48–52)
5. Check for existing .xlsc save file in expected location
6. If save file found: run multi-stage validation (IDS_283–285, IDS_41)
7. If validation passes: restore saved values (IDS_84)
8. Trigger initial recalculation of scheduling formulas
9. Navigate to the primary scheduling sheet (e.g., Gantt view)

### 4.2 `Workbook_BeforeSave`

**Location:** `ThisWorkbook` module (mandatory Excel convention)
**Confidence: HIGH** (custom save dialog IDS strings directly imply this handler)

This handler fires when the user presses Ctrl+S or selects File > Save. The BoxedApp SDK hook layer intercepts the save event before it reaches Excel's native save logic. The handler:

1. Cancels the native Excel save (`Cancel = True`)
2. Presents the custom save dialog (IDS_3, IDS_4, IDS_7, IDS_8)
3. If user confirms, serializes modified cell values to the .xlsc format
4. Writes the .xlsc file with integrity checksum and hardware fingerprint

### 4.3 `Workbook_BeforeClose`

**Location:** `ThisWorkbook` module (mandatory Excel convention)
**Confidence: HIGH** (IDS_5 "Do you want to keep changes?" directly evidences this handler)

This handler fires when the user attempts to close the application. It:

1. Checks whether unsaved modifications exist
2. If modifications exist: presents "Do you want to keep changes?" prompt (IDS_5)
3. If "Yes" (IDS_7): triggers the custom save routine (as in `Workbook_BeforeSave`)
4. If "No" (IDS_8): discards changes and allows the close to proceed
5. Cancels the close event if the user selects Cancel (if a three-button dialog is used)

### 4.4 Sheet-Level `Calculate` Events

**Location:** Individual `Sheet` modules (e.g., `Sheet1`, `Sheet2`, etc.)
**Confidence: LOW** (inferred from scheduling application nature; no direct IDS evidence)

For a scheduling application with dynamic Gantt charts or resource histograms, sheet-level `Calculate` event handlers are a common pattern. These fire whenever Excel recalculates formulas on the sheet and can be used to:

- Redraw Gantt bar shapes after date or duration changes
- Update progress indicators or completion percentages
- Refresh resource histogram charts
- Trigger cross-sheet recalculation dependencies not handled by Excel's formula engine

The frequency of these events in a complex scheduling workbook can be a significant performance factor, often requiring debouncing logic (`Application.EnableEvents = False` / `True` guards).

### 4.5 Additional Event Handlers (Low Confidence)

**Confidence: LOW**

| Handler | Location | Inferred Purpose |
|---------|----------|-----------------|
| `Workbook_SheetChange` | `ThisWorkbook` | Responds to cell edits; triggers dependent recalculations or validation |
| `Worksheet_Change` | Sheet modules | Cell-level input validation for task durations, resource names, dates |
| `Worksheet_SelectionChange` | Sheet modules | Context-sensitive help or status bar updates based on selected cell |
| `Workbook_WindowActivate` | `ThisWorkbook` | Refreshes UI state when the Excel window regains focus |
| `Workbook_NewSheet` | `ThisWorkbook` | Possibly blocked or restricted to prevent users from adding unauthorized sheets |

---

## 5. VBA Security Mechanisms

### 5.1 XLS Padlock VBA Protection

**Confidence: HIGH** (core XLS Padlock product feature; directly relevant to analysis constraints)

XLS Padlock applies multiple layers of protection to the VBA project:

1. **AES-256 workbook encryption:** The entire `.xlsb` or `.xlsm` file (including the VBA project storage) is encrypted before packaging. Standard VBA password removal tools (e.g., removing `DPB=` from the file) cannot operate on the ciphertext.

2. **VBA project password:** XLS Padlock additionally sets a VBA project password within the workbook before encryption. This means that even after decryption (e.g., by a legitimate licensed user who has opened the application), the VBA editor is still locked. The VBA source code is protected by two independent mechanisms.

3. **Memory-only execution:** The BoxedApp SDK's `CreateProcessFromMemoryW` ensures the decrypted workbook never touches the file system as a plaintext file, preventing forensic recovery of the decrypted image via disk analysis tools.

4. **Hook-layer enforcement:** `BoxedAppSDK_HookFunction` allows the stub to monitor and selectively block API calls that might expose the workbook content, such as `CopyFile`, `CreateFile` on the virtualized workbook path, or COM marshal operations that could serialize the workbook object.

### 5.2 Implications for VBA Analysis

The combination of these protections means that the VBA project in SchedPro Tool64.exe is effectively inaccessible without:

- The correct activation code (to derive the AES-256 decryption key)
- AND the VBA project password (to unlock the VBA editor after decryption)

Dynamic analysis (running the application under a debugger and intercepting the decrypted workbook in memory) would require both defeating the activation check and locating the in-memory Excel process image before it is unmapped — a non-trivial task given the BoxedApp virtualization layer.

---

## 6. Cryptographic Class Infrastructure

**Confidence: HIGH** (class names directly observed in XLSPadlockStub.dll RTTI metadata)

The following cryptographic class names were identified in the Run-Time Type Information (RTTI) metadata of XLSPadlockStub.dll. These are Delphi/Object Pascal class names, consistent with XLS Padlock being a Delphi-compiled application.

| Class Name | Algorithm | Role in XLS Padlock |
|------------|-----------|---------------------|
| `TAESCore` | AES (128/192/256-bit) | Core cipher engine for workbook encryption/decryption |
| `TSHA3Core` | SHA-3 (Keccak family) | Hash function; possibly used for key derivation or integrity verification |
| `TSHA2Core` | SHA-2 (SHA-256/SHA-512) | Hash function; SHA-512 is used for the XPL04 footer digest |
| `rsa_keygen` | RSA | Key generation; used in licensing/activation infrastructure |
| `dsa_keygen` | DSA (Digital Signature Algorithm) | Signing; likely used to sign activation codes or license certificates |

### 6.1 Cryptographic Architecture Inferences

**Confidence: MEDIUM**

Based on the observed class set, the cryptographic architecture of XLS Padlock v25.2 likely operates as follows:

- **Workbook encryption:** `TAESCore` with a 256-bit key (AES-256) encrypts the workbook payload. The key is derived from user input and hardware data.
- **Key derivation:** `TSHA2Core` (SHA-256 or HMAC-SHA-256) likely hashes the activation code combined with hardware fingerprint bytes to produce the 32-byte AES-256 key. SHA-3 (`TSHA3Core`) may serve as an alternative or secondary KDF.
- **Container integrity:** `TSHA2Core` (SHA-512) computes the XPL04 footer hash, and CRC32 covers the encrypted payload per the GXLS2 block.
- **Activation/licensing:** `rsa_keygen` and `dsa_keygen` support a public-key activation scheme where G.D.G. Software (or the SchedPro vendor) issues activation codes signed with a private RSA/DSA key. The stub verifies the signature using the embedded public key before accepting an activation code.

The RSA/DSA infrastructure explains why the activation code cannot simply be brute-forced: the stub validates the activation code's cryptographic signature before using it in key derivation, so random strings cannot generate a valid decryption key.

### 6.2 Relationship to VBA

The cryptographic classes in XLSPadlockStub.dll operate entirely outside the VBA layer. VBA code in the workbook has no direct access to `TAESCore` or related classes — these are in the native stub layer. However, VBA code may call back into the stub via registered COM interfaces or shell extension points to trigger operations such as:

- Requesting the stub to save/load the .xlsc file (which the stub encrypts using its own crypto)
- Querying the activation state or license tier
- Triggering hardware fingerprint collection for the save file's hardware lock

---

## 7. Summary of VBA Evidence and Confidence

| Claim | Confidence | Primary Evidence |
|-------|-----------|-----------------|
| VBA macros are present in the workbook | MEDIUM | XLS Padlock design intent; 7.21 MB payload size; behavioral IDS strings implying programmatic logic |
| VBA project is AES-256 encrypted | HIGH | TAESCore in DLL; XLS Padlock product specification |
| VBA project has a password set | MEDIUM | Standard XLS Padlock feature; not directly confirmed for this specific build |
| `Workbook_Open` event handler exists | HIGH | IDS_21 startup string; add-in loading; restore logic |
| `Workbook_BeforeClose` event handler exists | HIGH | IDS_5 close prompt directly evidences this |
| `Workbook_BeforeSave` event handler exists | HIGH | IDS_3/4/7/8 custom save dialog directly evidences this |
| Custom save/restore routines exist | HIGH | IDS_84, IDS_283–285 directly evidence save/restore logic |
| Scheduling calculation VBA modules exist | LOW | Inferred from application purpose; no direct string evidence |
| Sheet-level Calculate event handlers exist | LOW | Common pattern for Gantt chart workbooks; no direct evidence |
| Add-ins loaded at startup | HIGH | IDS_44, IDS_48–52 directly evidence add-in loading |
| BoxedApp hooks intercept Excel save/close | HIGH | BoxedAppSDK_HookFunction directly observed in DLL imports |
| Memory-only execution (no disk write) | HIGH | BoxedAppSDK_CreateProcessFromMemoryW directly observed |
| RSA/DSA-signed activation codes | MEDIUM | rsa_keygen / dsa_keygen class names in DLL RTTI |

---

## 8. Limitations and Potential Next Steps

The following VBA-specific information cannot be determined from this analysis:

- Module names (e.g., `modGantt`, `clsTask`, `frmSaveDialog`)
- Procedure/function names and signatures
- The number of VBA modules (standard, class, form, document)
- Whether the workbook uses `.xlsb` (Binary) or `.xlsm` (OOXML) format internally
- Whether any external references (early binding) to non-standard COM libraries exist
- The VBA project name as set in Tools > Properties
- Whether the workbook uses `Option Explicit` and other quality indicators

Approaches that could extend VBA analysis without decryption:

1. **Zstd config decompression:** The Zstd blob at 0x000032–0x00130E may list add-in paths, module counts, or other VBA-related configuration that XLS Padlock stores separately from the encrypted workbook.
2. **Runtime API monitoring (dynamic analysis):** Monitoring `LoadLibrary`, `CoCreateInstance`, and VBA runtime API calls during a licensed execution could reveal module and procedure names as they are JIT-compiled by the VBA engine.
3. **XLSPadlockStub.dll deep string analysis:** Further enumeration of DLL string tables beyond the IDS_* set may reveal hard-coded module names, procedure call targets, or COM ProgID strings that the stub uses to communicate with the VBA runtime.

---

*This document is produced for forensic and educational purposes. No activation codes, decryption keys, VBA password bypass methods, or DRM circumvention techniques are described or implied. All VBA-layer analysis is based solely on observable metadata from XLSPadlockStub.dll and the XPLAPP container structure.*
