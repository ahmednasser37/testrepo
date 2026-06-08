# SchedPro Tool – Business Logic Reconstruction

**Target:** SchedPro_Tool64.exe  
**Classification:** Behavioral/Structural Analysis (Encrypted Workbook)  
**Date:** 2026-06-08  
**Analyst Note:** This document is a forensic/educational analysis for software understanding purposes only. All claims are derived from string artefacts, executable metadata, and known XLS Padlock platform patterns. No instructions for bypassing DRM, licensing, or authentication are provided or implied.

---

## 1. Executive Summary

SchedPro Tool is a **commercial Windows desktop project scheduling application** built on Microsoft Excel and distributed via **XLS Padlock v25.2**. The application packages an AES-256 encrypted Excel workbook inside a 64-bit native Windows executable (`SchedPro_Tool64.exe`). At runtime, XLS Padlock's stub library (`XLSPadlockStub.dll`) decrypts the workbook in memory, launches a hidden Excel host process, and presents the resulting spreadsheet-based scheduling tool to the user as a standalone desktop application.

The product targets **professional project managers and schedulers** — primarily in European markets — who require a capable scheduling tool without the cost or complexity of dedicated scheduling platforms. It competes in the Excel-based scheduling tool segment against products such as Microsoft Project (light-end overlap), Primavera P6 (high-end overlap), and other Excel-based scheduling add-ins.

---

## 2. Problem Domain

**Domain:** Professional project scheduling and management.

Project scheduling tools address the following core problems:

- Defining and organising project tasks in a hierarchical Work Breakdown Structure (WBS).
- Establishing task dependencies (finish-to-start, start-to-start, etc.) and computing the resulting schedule network.
- Identifying the critical path — the sequence of tasks that determines the minimum project duration.
- Tracking progress (percent complete, actual start/finish) against a frozen baseline plan.
- Managing resource assignments, availability, and cost.
- Producing reports and visualisations (Gantt charts, S-curves, resource histograms) for stakeholder communication.

SchedPro Tool addresses this domain using Excel as the calculation and display platform, with VBA macros providing the interactive and dynamic behaviour layer.

---

## 3. Target Users

**Primary:** Project managers, project schedulers, and planning engineers who:

- Work in project-delivery industries (construction, engineering, IT, manufacturing).
- Are based in or serve European markets (implied by `dd/mm/yyyy` date format in `WebUpdate1.DateFormat`).
- Use Microsoft Excel as part of their daily workflow.
- Require a scheduling tool that operates as a familiar, Excel-native environment.
- May not have access to or budget for enterprise platforms such as Oracle Primavera P6.

**Secondary:** Organisations distributing a standardised scheduling template to project teams, leveraging XLS Padlock's copy-protection to control distribution of proprietary scheduling frameworks.

**Confidence:** Medium – inferred from date locale, application class, and XLS Padlock's typical commercial use cases.

---

## 4. Application Lifecycle

The following sequence describes the normal operating lifecycle of SchedPro Tool, reconstructed from stub string artefacts and XLS Padlock architectural documentation.

### Step 1 — Launch

`SchedPro_Tool64.exe` is executed by the user. The Windows loader initialises the process; `XLSPadlockStub.dll` is loaded and takes control of startup. The stub performs environment checks:

- Verifies that a compatible version of Microsoft Excel (desktop, 64-bit) is installed.
- Confirms that Excel is not the Microsoft Store (AppX) edition, which is unsupported (see §8).
- Checks for an existing running instance (single-instance enforcement, `IDS_81`).
- Verifies the process is not running with elevated (administrator) privileges (`IDS_79`).

### Step 2 — License Verification

The stub evaluates the license state before proceeding. License modes include:

- **Activated:** A valid software activation key is present and bound to the current hardware system ID.
- **Dongle:** A USB hardware security dongle is present and valid.
- **Trial:** No activation present; trial run count and/or day limit applies.
- **Expired/Invalid:** Trial exhausted or activation invalid; launch is blocked.

If the license check passes, the stub proceeds to the welcome screen.

### Step 3 — Welcome Screen

The user is presented with a choice:

- **Open Original Workbook:** Load the master template in its factory state. No prior user data is loaded.
- **Open Save File (.xlsc):** Load a previously saved user session. The stub validates the save file's hardware binding and integrity checksum before proceeding.

Recent save files are listed for quick access (recent-files list maintained by the stub).

### Step 4 — Excel Launch

Excel is launched as a hidden or borderless host process. The stub:

1. Decrypts the AES-256 workbook in memory.
2. Loads the decrypted workbook into Excel.
3. If a `.xlsc` save file was selected, injects previously saved values into the workbook ("Restoring values previously saved", `IDS_84`).
4. Enables required Excel add-ins.
5. Presents the Excel window to the user, styled as the SchedPro Tool application.

### Step 5 — Scheduling Work

The user interacts with the Excel-based scheduling tool:

- Entering and editing project tasks, durations, and dependencies.
- Excel's calculation engine recalculates the schedule model in real time.
- VBA macros respond to user actions (data validation, UI navigation, report generation).
- The user remains within the Excel environment; the application is indistinguishable from a native Windows application.

### Step 6 — Save / Close

When the user closes the application or requests a save:

- The stub captures the current workbook state (modified cell values, user inputs).
- The user is prompted to save as a `.xlsc` file (hardware-locked to the current machine).
- If the user declines, changes are discarded and the master template is unmodified.
- The Excel host process is terminated; the stub exits.

**Confidence:** High for overall flow (XLS Padlock platform invariant); Medium for specific UI/dialog details (inferred from IDS strings).

---

## 5. License and DRM Model

SchedPro Tool employs XLS Padlock's full DRM feature set. The following mechanisms are in use, evidenced by stub string artefacts (`IDS_0` through `IDS_115` and beyond).

### 5.1 Software Activation Key

The primary licensing mechanism. A unique alphanumeric activation key is issued to each licensed user. The key is bound to the **hardware system ID** of the target machine (computed from CPU, motherboard, and other stable hardware identifiers). Activation may occur:

- **Online:** The stub contacts an activation server; the key-hardware binding is validated and recorded server-side.
- **Manual (offline):** The user provides a system ID to the software vendor, who issues an activation certificate that can be applied without internet access.

### 5.2 USB Hardware Dongle

An optional secondary licensing mechanism. A USB security dongle (hardware token) is queried at startup. If present and valid, the dongle alone is sufficient for licensing regardless of the machine's hardware ID. This mode supports users who work on multiple machines (e.g., office desktop and site laptop) and carry the dongle with them.

### 5.3 Trial Mode

Unlicensed installations enter trial mode, governed by:

- **Run count limit:** `IDS_32` — "You have %u %s left..." — a countdown of remaining trial launches.
- **Day limit:** The trial period may also be bounded by a fixed number of calendar days from first launch.

When the trial is exhausted, the application blocks further launches until a valid activation key is provided.

### 5.4 License Transfer (Deactivation)

A user may transfer their license to a new machine by **deactivating** on the current machine. Deactivation produces a certificate that the activation server records, freeing the activation slot for re-use on a different hardware ID. Manual deactivation certificates are supported for machines without internet access.

### 5.5 Hardware-Locked Save Files

`.xlsc` save files are cryptographically bound to the hardware ID of the machine on which they were created. Attempting to open a `.xlsc` file on a different machine produces the error: "hardware-locked and cannot be opened on this computer" (`IDS_41`). This prevents sharing of licensed save files across machines as an alternative to purchasing additional licenses.

**Confidence:** High – all mechanisms are documented in XLS Padlock platform capabilities and corroborated by IDS string artefacts.

---

## 6. Data Model

SchedPro Tool maintains three logical data layers:

### 6.1 Original Workbook (Master Template)

The AES-256 encrypted workbook embedded in the executable. This is the **factory-state template**: it contains all formulas, VBA code, formatting, and structural definitions but no user project data. It is never modified by user activity; it serves as the starting point for every new project.

**Confidence:** High – invariant property of the XLS Padlock architecture.

### 6.2 Save Files (.xlsc)

User-generated snapshots of the workbook state after data entry and scheduling work. A `.xlsc` file contains:

- The delta of user-entered cell values relative to the master template.
- Computed scheduling values (dates, floats, EVM metrics) as of the last save ("Restoring values previously saved", `IDS_84`).
- Hardware-binding metadata linking the file to a specific machine.
- Integrity checksums validated on load (`IDS_284`, `IDS_285`).

The `.xlsc` format is proprietary to XLS Padlock; it is not a standard Excel file and cannot be opened directly in Excel.

**Confidence:** High – XLS Padlock platform invariant corroborated by IDS artefacts.

### 6.3 Recent Save File History

The stub maintains a list of recently accessed `.xlsc` files, presented on the welcome screen for quick access. This is analogous to the "Recent Documents" list in standard desktop applications.

**Confidence:** Medium – standard XLS Padlock feature; inferred from platform documentation.

### 6.4 Save Integrity Validation

On load, each `.xlsc` file is subject to:

- **Hardware binding check:** The file's bound hardware ID must match the current machine.
- **Integrity checksum:** The file contents are hashed and compared to a stored checksum. A mismatch indicates disk corruption or file tampering and produces an "Incorrect save" or "Corrupted save" error (`IDS_284`, `IDS_285`).

**Confidence:** High – artefact-sourced.

---

## 7. Feature Set

The following features are inferred from IDS string artefacts and platform capabilities.

| Feature | Evidence | Confidence |
|---|---|---|
| Project scheduling (core) | Application name, domain context | High |
| Save / restore workflow | `IDS_84`, `.xlsc` format | High |
| Multiple save files with recent history | Welcome screen recent-files list | Medium |
| Single-instance enforcement | `IDS_81` | High |
| Copy protection (cell/data) | `IDS_18`: "Copy is not allowed" | High |
| Print protection (optional DRM) | `IDS_34`: "printing is not allowed" | High |
| Software activation (online/offline) | `IDS_0`–`IDS_115` activation strings | High |
| USB dongle support | Dongle IDS strings | High |
| Trial mode with countdown | `IDS_32`: "%u %s left" | High |
| License deactivation / transfer | Deactivation IDS strings | High |
| Hardware-locked save files | `IDS_41` | High |
| WebUpdate (auto-update) | `WebUpdate1.DateFormat` config key | High |
| Excel add-in integration | Add-in check IDS strings | Medium |
| Multi-language support (English primary) | String table structure | Medium |
| Gantt chart visualisation | Scheduling tool domain norm | Low |
| Resource management | Application class and size | Low |
| Earned value reporting | Professional scheduling tool class | Low |

---

## 8. Key Business Rules

The following business rules are enforced by the application at runtime:

### 8.1 Single Instance Only

Only one running instance of SchedPro Tool is permitted per machine at any time (`IDS_81`). A second launch attempt will detect the existing instance and either bring it to focus or display an error.

**Confidence:** High – artefact-sourced.

### 8.2 Save Files Are Machine-Locked

`.xlsc` save files are cryptographically bound to the hardware ID of the machine that created them and cannot be opened on any other machine (`IDS_41`). This rule is enforced at file-load time and cannot be overridden by the user.

**Confidence:** High – artefact-sourced.

### 8.3 Desktop Excel Required

The application requires Microsoft Excel to be installed as a standard desktop application. The Microsoft Store (AppX) edition of Excel is explicitly not supported. This rule exists because XLS Padlock injects a DLL into the Excel process; AppX sandboxing prevents this injection.

**Confidence:** High – standard XLS Padlock constraint corroborated by IDS strings.

### 8.4 Normal User Privileges Required

The application must be run as a **standard (non-administrator) Windows user** (`IDS_79`). Running with elevated privileges is blocked. This is a security measure to prevent the application from operating in an environment that could be used to circumvent DRM mechanisms (e.g., kernel-level debugging tools accessible only to administrators).

**Confidence:** High – artefact-sourced.

### 8.5 Add-Ins Must Be Enabled

Required Excel add-ins must be enabled for core scheduling functionality to operate. If add-in loading fails, the application will display an error and may refuse to proceed. This reflects the application's dependency on specific Excel extensibility hooks.

**Confidence:** Medium – inferred from add-in check IDS strings.

### 8.6 Copy and Print Restrictions

Data copying (`IDS_18`: "Copy is not allowed") and printing (`IDS_34`: "printing is not allowed") may be restricted by the publisher's DRM configuration. These restrictions are enforced via VBA event handlers that intercept Excel's copy and print actions. The publisher may enable or disable these restrictions per distribution.

**Confidence:** High – artefact-sourced; specific enablement state unknown.

---

## 9. Competitive Context

SchedPro Tool occupies the **Excel-based professional scheduling** segment, which sits between lightweight spreadsheet-only approaches and full enterprise scheduling platforms:

| Tier | Products | Comparison |
|---|---|---|
| Enterprise | Oracle Primavera P6, Safran Risk | Full CPM engine, P6-compatible XER/XML, dedicated DB backend; far higher cost and complexity |
| Mid-market | Microsoft Project (MSP) | Native CPM engine, Gantt, resource leveling; requires MSP license; not Excel-native |
| Excel-based | SchedPro Tool, Asta Powerproject (Excel import), custom templates | Excel-native, lower cost, familiar UX; calculation engine limited to Excel's capabilities |
| Free/open | OpenProject, ProjectLibre | Open-source; no commercial DRM; different UX model |

**Positioning:** SchedPro Tool targets users who want the familiarity and flexibility of Excel with a professionally structured scheduling framework, protected by commercial DRM to allow monetisation of the proprietary scheduling template and methodology.

**Confidence:** Medium – competitive analysis is domain-knowledge-based; no artefact directly references competitors.

---

## 10. Confidence Assessment Summary

| Claim | Confidence |
|---|---|
| Commercial Windows desktop application | High |
| Built on Microsoft Excel via XLS Padlock v25.2 | High |
| AES-256 encrypted workbook | High |
| 64-bit process architecture | High |
| European (dd/mm/yyyy) date locale | High |
| Professional project scheduling domain | High |
| Software activation key (hardware-bound) | High |
| USB dongle support | High |
| Trial mode with run/day countdown | High |
| Hardware-locked .xlsc save files | High |
| Save integrity checksums | High |
| Single-instance enforcement | High |
| Copy/print restriction capability | High |
| WebUpdate auto-update mechanism | High |
| Normal-user privilege requirement | High |
| Desktop Excel (non-AppX) requirement | High |
| Target users: project managers / schedulers | Medium |
| European / international primary market | Medium |
| Complex multi-sheet workbook structure | Medium |
| CPM algorithm implementation | Medium |
| WBS hierarchy support | Medium |
| Baseline vs. actual tracking | Medium |
| Add-in dependency | Medium |
| Recent save files list | Medium |
| Gantt chart visualisation | Low |
| Resource loading and leveling | Low |
| Earned value management (EVM) | Low |
| Competitive positioning vs. MSP / P6 | Medium |

---

*This document contains forensic analysis findings for software understanding and educational purposes only. No information herein constitutes instructions for bypassing commercial licensing, DRM, or authentication systems.*
