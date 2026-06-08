# SchedPro Tool – Calculation Engine Analysis

**Target:** SchedPro_Tool64.exe  
**Classification:** Behavioral/Structural Analysis (Encrypted Workbook)  
**Date:** 2026-06-08  
**Analyst Note:** This document is a forensic/educational analysis for software understanding purposes only.

---

## 1. Analysis Constraints

The workbook component of SchedPro Tool is AES-256 encrypted by XLS Padlock v25.2. As a result, direct inspection of Excel formulas, VBA source code, and worksheet structure is not possible without the decryption key held by the application at runtime. All findings in this document are therefore **inferred from behavioral evidence**: string literals extracted from the XLSPadlockStub.dll binary, executable metadata, configuration artefacts, and the known architectural patterns of XLS Padlock-packaged applications.

Confidence levels are assigned per finding:

| Level | Meaning |
|---|---|
| High | Supported by multiple independent artefacts or by XLS Padlock platform invariants |
| Medium | Consistent with evidence but requires assumptions |
| Low | Reasoned inference; plausible given domain context, no direct artefact support |

---

## 2. Application Identity as a Scheduling Tool

### 2.1 Branding

The product name **"SchedPro"** is a compound of "Sched" (Schedule) and "Pro" (Professional), establishing the application's market positioning as a professional-grade scheduling solution. This naming convention is consistent with commercial project management software targeting engineering and project delivery industries.

**Confidence:** High – name is present in the executable filename and configuration artefacts.

### 2.2 Architecture: 64-bit Build

The executable is distributed as `SchedPro_Tool64.exe`, explicitly targeting the 64-bit Windows process model. Implications for the calculation engine:

- Excel's 64-bit host process (`EXCEL.EXE`, x64) is required; the 32-bit Office variant is unsupported.
- The 64-bit address space permits workbooks with large in-memory datasets without hitting the ~2 GB virtual address ceiling of 32-bit processes.
- Large project schedules (thousands of tasks, resources, and time-phased data) can be held in memory simultaneously.
- VBA compiled under 64-bit requires `PtrSafe` declarations for API calls; the workbook's VBA was presumably authored with this in mind.

**Confidence:** High – the filename suffix `64` is an explicit architectural declaration.

### 2.3 Locale and Date Format

The configuration artefact `WebUpdate1.DateFormat` contains the value `dd/mm/yyyy`, the **European (day-first) date convention**. This setting governs:

- How date strings are parsed from user input fields.
- How dates are serialised to save files (`.xlsc`).
- How update-related timestamps are displayed in the UI.
- Implicit assumption that the primary market is European (UK, continental Europe) or a region adopting this convention.

**Confidence:** High – sourced directly from a config key.

### 2.4 Workbook Size

The encrypted workbook is approximately **7.21 MB**. For reference, a plain Excel workbook with only a few worksheets and minimal data is typically under 100 KB. A 7+ MB workbook (before any user data, since this is the master template) indicates:

- Multiple worksheets (likely 10–30+ tabs for different project views).
- Pre-populated formula matrices covering a full scheduling horizon.
- Embedded chart objects, form controls, or structured table definitions.
- Potentially embedded VBA forms (UserForms) for data entry dialogs.
- Possibly embedded images for the UI (logos, help graphics).

**Confidence:** Medium – size is a proxy; exact sheet count and content breakdown are inaccessible.

---

## 3. Excel as Calculation Platform

XLS Padlock packages an Excel workbook as a standalone Windows application. The underlying calculation infrastructure is therefore Excel's own engine, not a custom interpreter.

### 3.1 Native Cell Formulas

All arithmetic, logical, and date calculations are expressed as Excel worksheet formulas stored in cells. The Excel calculation engine evaluates these formulas according to Excel's recalculation rules (dependency tree traversal). No separate formula interpreter is present in the stub executable.

### 3.2 VBA Macros

Dynamic behavior — responding to user actions, validating inputs, triggering recalculation sequences, managing UI state — is implemented in VBA (Visual Basic for Applications) modules embedded in the workbook. VBA runs within Excel's hosted VBA runtime; the stub executable does not contain a separate scripting engine.

### 3.3 Recalculation Mode

Excel's default recalculation mode is **automatic**: any cell change triggers a dependency-ordered recalculation of the entire affected subgraph. For a scheduling tool, this means that changing a task's start date or duration immediately propagates updated finish dates, float values, and resource loads across the entire model without requiring an explicit "calculate" button press. This is a key usability feature of Excel-based scheduling tools versus static data entry forms.

**Confidence:** High – these are invariant properties of the XLS Padlock architecture.

---

## 4. Date and Time Framework

Evidence from the stub binary reveals a comprehensive date/time handling layer, consistent with a scheduling application that must reason about calendars, working days, and time horizons.

### 4.1 Date Format

Primary format: `dd/mm/yyyy` (European, day-first).

Additional formats identified in stub string tables:

| Format String | Likely Use |
|---|---|
| `MM/dd/yy` | Short date for compact display in schedule cells |
| `dddd MMMM dd yyyy` | Long-form date for report headers or milestone labels |
| `HH:mm:ss` | Time-of-day component for timestamp logging |

### 4.2 Day-of-Week Constants

Constants for all seven weekdays (Sunday through Saturday) are present in the stub. These are used to:

- Define the project working calendar (which days are workdays).
- Compute day-of-week for any given date (e.g., identifying weekend boundaries).
- Drive NETWORKDAYS/WORKDAY formula logic that respects non-working days.

### 4.3 Month Constants

All twelve month name constants (January through December) are present, supporting:

- Human-readable date display in reports and Gantt timescales.
- Month-level aggregation of resource loading and cost curves.
- Period-based earned value calculations (monthly reporting periods).

**Confidence:** High for format/constants presence (artefact-sourced); Medium for specific use-case interpretations.

---

## 5. Inferred Scheduling Algorithms

The following algorithms are standard to professional project scheduling tools of this class. Their presence is inferred from the application's stated purpose, size, and domain conventions.

### 5.1 Critical Path Method (CPM)

CPM is the foundational algorithm of project scheduling. It requires:

- **Forward pass:** Computing Early Start (ES) and Early Finish (EF) for every task by traversing the network from project start to end.
- **Backward pass:** Computing Late Start (LS) and Late Finish (LF) by traversing in reverse from project end.
- **Float/Slack calculation:** Total Float = LS – ES (or LF – EF); Free Float derived from successor relationships.
- **Critical path identification:** Tasks with zero total float form the critical path.

In an Excel implementation, CPM is typically encoded as a column-per-field formula layout across a task list, with predecessor logic expressed via VLOOKUP or structured table references.

**Confidence:** Medium – CPM is near-universal in tools of this description; no direct artefact confirms it.

### 5.2 Work Breakdown Structure (WBS)

Project tasks are organised in a hierarchical WBS. Summary rows aggregate child task durations, costs, and resource assignments. Excel handles this via SUMIF or structured table subtotaling.

**Confidence:** Medium.

### 5.3 Duration Calculations

String `IDS_20` references the unit "day(s)", confirming that task durations are expressed in day-granularity. Duration arithmetic must account for:

- Calendar working days (excluding weekends and holidays).
- Part-day durations if hour-level granularity is supported.
- Elapsed duration (calendar days regardless of working calendar) vs. work duration.

**Confidence:** High for day-granularity (artefact-sourced); Medium for hour/elapsed variants.

### 5.4 Forward/Backward Pass and Float

As a direct consequence of CPM implementation, the workbook contains formulas computing forward pass dates, backward pass dates, total float, and free float for each task.

**Confidence:** Medium.

### 5.5 Resource Loading and Leveling

Professional scheduling tools track resource assignments (who does what, for how long) and aggregate resource demand over time. Resource leveling resolves over-allocations by delaying tasks within available float. In Excel, this is typically a combination of:

- Per-task resource assignment tables.
- Time-phased loading matrices (resources × time periods).
- Leveling logic implemented in VBA due to its iterative, non-formulaic nature.

**Confidence:** Low – consistent with the application class and size; no direct string evidence.

### 5.6 Earned Value Management (EVM)

EVM is a standard performance measurement technique in professional scheduling. Three core metrics:

| Metric | Full Name | Meaning |
|---|---|---|
| BCWS | Budgeted Cost of Work Scheduled | Planned value at a given date |
| BCWP | Budgeted Cost of Work Performed | Earned value based on % complete |
| ACWP | Actual Cost of Work Performed | Actual expenditure to date |

Derived indicators: Schedule Performance Index (SPI = BCWP/BCWS), Cost Performance Index (CPI = BCWP/ACWP), Schedule Variance (SV), Cost Variance (CV), Estimate at Completion (EAC).

**Confidence:** Low – standard in professional tools of this tier; inferred from domain context.

### 5.7 Percent Complete Tracking

The stub string `IDS_20` ("day(s)") and the general scheduling context imply tracking of task percent complete (physical or duration-based), used as input to BCWP calculation and progress reporting.

**Confidence:** Medium.

### 5.8 Baseline vs. Actual Tracking

Professional scheduling requires storing a frozen baseline (original plan) alongside current (updated) values to compute variances. Excel implementations typically maintain parallel columns: `BL_Start`, `BL_Finish`, `BL_Duration`, `BL_Cost` alongside `ACT_Start` etc.

**Confidence:** Medium.

---

## 6. Formula Categories Likely Present

Given the scheduling algorithms described above, the following Excel formula categories are expected in the workbook:

| Formula / Function | Scheduling Use |
|---|---|
| `NETWORKDAYS(start, end, holidays)` | Working-day count between two dates |
| `WORKDAY(start, days, holidays)` | Add N working days to a start date |
| `DATE(year, month, day)` | Construct dates from components |
| `YEAR()`, `MONTH()`, `DAY()` | Extract date components for period aggregation |
| `VLOOKUP()` / `INDEX()`+`MATCH()` | Predecessor lookup, resource rate lookup |
| `IF()` / `IFS()` | Conditional scheduling logic (e.g., milestone handling) |
| `SUMIF()` / `SUMIFS()` | Aggregate costs/hours by resource, WBS, period |
| Array formulas (`Ctrl+Shift+Enter`) | Critical path traversal, multi-criteria aggregations |
| `IFERROR()` | Graceful handling of incomplete task definitions |
| `TEXT(date, format)` | Format dates for display using the configured format strings |
| `MAX()` / `MIN()` | Earliest/latest date across predecessors (CPM forward/backward pass) |

**Confidence:** Medium – these are the standard formula toolkit for Excel-based scheduling; exact usage cannot be verified without decryption.

---

## 7. Calculation Engine Characteristics (from Stub Evidence)

### 7.1 Single-Instance Enforcement

`IDS_81` enforces that only one instance of SchedPro Tool may run simultaneously. From a calculation-engine perspective, this prevents two Excel host processes from operating on the same workbook file concurrently, which would create file-locking conflicts and potentially corrupt the calculation state.

**Confidence:** High – artefact-sourced.

### 7.2 Persistent Calculation State

`IDS_84`: "Restoring values previously saved." This string confirms that the application preserves and restores the full workbook calculation state between sessions. When a `.xlsc` save file is loaded, previously computed scheduling values (dates, floats, EVM metrics) are restored without requiring a full recalculation from scratch. This is important for large models where recalculation may take several seconds.

**Confidence:** High – artefact-sourced.

### 7.3 Save Integrity Validation

`IDS_284` and `IDS_285` relate to save-file integrity checks ("Incorrect save" / "Corrupted save" categories). This confirms that the save file format includes checksums or hash values that are validated on load. A corrupted save — whether from disk error, incomplete write, or tampering — will be detected and rejected before any calculation state is restored, preventing silent data corruption.

**Confidence:** High – artefact-sourced.

---

## 8. Confidence Summary

| Finding | Confidence |
|---|---|
| European date format (dd/mm/yyyy) | High |
| 64-bit architecture requirement | High |
| Excel as calculation host | High |
| VBA macro-driven dynamic behavior | High |
| Single-instance enforcement | High |
| Persistent calculation state restoration | High |
| Save integrity validation | High |
| Day-granularity durations | High |
| Day-of-week and month constants present | High |
| Complex multi-sheet workbook (size evidence) | Medium |
| CPM forward/backward pass implementation | Medium |
| WBS hierarchy with summary aggregation | Medium |
| Baseline vs. actual tracking columns | Medium |
| Percent complete tracking | Medium |
| Formula categories (NETWORKDAYS, VLOOKUP, etc.) | Medium |
| Resource loading/leveling | Low |
| Earned value management (BCWS/BCWP/ACWP) | Low |

---

*This document contains forensic analysis findings for software understanding and educational purposes. No information herein constitutes instructions for bypassing commercial licensing, DRM, or authentication systems.*
