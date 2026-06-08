# SchedPro Tool — User Interface Map

**Analysis Type:** Static forensic reconstruction from DFM binary form resources and IDS_* string table  
**Source Files:** XLSPadlockStub.dll (DFM resources), xlspadlock_decompressed.bin (UI strings)  
**Confidence Level:** Medium — form component hierarchies extracted from binary DFM resources; button labels and messages from string table; exact layout positions unknown due to DFM binary encoding complexity  
**Ethics Note:** This document is produced for educational, defensive, and interoperability purposes. No licensing bypass instructions are included.

---

## 1. Overview

SchedPro Tool's user interface consists of two layers:

1. **XLS Padlock Runtime Layer** — a Delphi VCL application (XLSPadlockStub.dll) providing DRM/activation dialogs, welcome screen, and file management UI before the workbook loads.
2. **Excel Workbook Layer** — the actual scheduling application UI rendered by Microsoft Excel after the workbook is decrypted and loaded in memory. This layer is inaccessible without activation.

This document covers all forms and dialogs in the XLS Padlock runtime layer, which are fully extractable from the binary.

---

## 2. Form Inventory

| Form Class       | Size (bytes) | Description                        | Font     |
|------------------|--------------|------------------------------------|----------|
| `TFWELC`         | 5,564        | Welcome / startup chooser          | Tahoma   |
| `TFACTKEY`       | 5,011        | Activation key entry               | (default)|
| `TFONACT`        | 5,432        | Online activation                  | (default)|
| `TFONVAL`        | 2,659        | Online validation                  | (default)|
| `TFDEACTKEY`     | 7,773        | Deactivation / license transfer    | (default)|
| `TFEULA`         | 3,895        | EULA / End User License Agreement  | (default)|
| `TTMSLOGINFORM`  | 21,162       | TMS Software login form            | (default)|
| `TWUWIZ`         | 74,055       | WebUpdate wizard                   | Verdana  |
| `TDATAMODULE1`   | 35,979       | Data module (non-visual)           | N/A      |

---

## 3. Form Descriptions

### 3.1 Welcome Screen (`TFWELC`)

**Trigger:** Appears on application startup after license verification passes.  
**Caption:** `Welcome`  
**Window Style:** `bsDialog` (modal dialog, non-resizable)  
**Position:** `poScreenCenter` (centered on screen)  
**Events:** `OnCreate → FormCreate`, `OnDestroy → FormDestroy`  
**Font:** Tahoma  
**Layout Components:**
- `TBevel` (Bevel1) — decorative top border line (`bsTopLine`), aligned to top
- `THTMLForm` label — HTML-rendered prompt label, aligned to top with margins (Left, Top, Right set; Width=88, Height=74)

**Displayed Text (from IDS_45, IDS_47):**
> Welcome! Choose what you want to do:

> Click "Original Workbook" if you do not have any previous workbook save file. Otherwise, you can select an existing secure workbook file to load.

**Buttons (inferred from IDS strings):**
| Button Label          | IDS Key | Action                                    |
|-----------------------|---------|-------------------------------------------|
| Original Workbook     | IDS_9   | Loads the original (unmodified) workbook  |
| Load A Recent Save    | IDS_12  | Shows list of previously saved workbooks  |
| Choose Save File      | IDS_14  | Opens file browser for `.xlsc` files      |

**Additional behavior:**
- IDS_46: "Excel is busy. Please verify if some dialog box is open... wait some seconds and click this button" — a Refresh button appears when Excel is slow to respond
- IDS_58: "Click to reload protected workbook" — a reload trigger button
- IDS_70: "If nothing happens, click this button to refresh..." — fallback refresh button

---

### 3.2 Activation Key Form (`TFACTKEY`)

**Trigger:** First run or when activation has expired/is missing.  
**Caption:** (not visible in DFM binary fragment, likely "Activation")  
**Window Style:** `bsDialog`  

**Displayed Text (from IDS_24, IDS_25):**
> To access [AppTitle], please enter your activation key and click **Activate**.

> You will have to provide the following system ID in order to receive an activation key:

**Fields:**
- System ID display (read-only, shows hardware fingerprint)
- Activation key input field

**Buttons:**
| Button Label              | IDS Key | Action                                    |
|---------------------------|---------|-------------------------------------------|
| Activate                  | —       | Submits key to validation logic           |
| Enter Other Activation Key| IDS_87  | Allows entering an alternate key          |
| Purchase Online           | IDS_39  | Opens browser to purchase page            |

**Error Messages:**
- IDS_38: `The key is not valid. Cannot continue.`
- IDS_42: `Invalid key.`
- IDS_23: `Your activation key has expired. Please provide a new one.`
- IDS_76: `This key has already been provided and/or deactivated.`
- IDS_78: `Please fill in mandatory fields (%s)`

---

### 3.3 Online Activation Form (`TFONACT`)

**Trigger:** When user chooses online activation method.  
**Window Style:** `bsDialog`  
**Size:** 5,432 bytes

**Purpose:** Sends activation key + system ID to XLS Padlock activation server (xlspadlock.com) and receives validation response.

**Displayed Text:**
- IDS_37: `[AppTitle] has been successfully activated with your key.`

**Error States:**
- Network timeout or server error handling
- IDS_86: `Failed Validation`

---

### 3.4 Online Validation Form (`TFONVAL`)

**Trigger:** Periodic online license revalidation.  
**Window Style:** `bsDialog`  
**Size:** 2,659 bytes (smallest form, simple confirm/cancel)

**Error Messages:**
- IDS_7 (form-specific): `An error occurred while reading server response:`
- IDS_8 (form-specific): `Are you sure you want to cancel validation? The application will exit.`
- IDS_9 (form-specific): `Invalid server response format`

---

### 3.5 Deactivation Form (`TFDEACTKEY`)

**Trigger:** User requests to transfer license to another machine.  
**Window Style:** `bsDialog`  
**Size:** 7,773 bytes

**Displayed Text (from form-specific IDS strings):**
> To transfer your [AppTitle] license to a new computer, the application must be deactivated on the current machine. This action will generate a Deactivation Certificate. It is important to copy or save this certificate as a text file, as it must be submitted to the software publisher for license transfer. Please be aware that once deactivated, the software will no longer be operational on this computer.

> To transfer [AppTitle] to another computer, please deactivate the application from this computer first. The activation server will be contacted automatically. If you do not have an Internet connection, you can choose Manual Deactivation.

**Buttons:**
| Button Label        | IDS Key | Action                                    |
|---------------------|---------|-------------------------------------------|
| Deactivate          | IDS_282 | Initiates online deactivation             |
| Manual Deactivation | —       | Generates deactivation certificate offline|
| Close               | IDS_22  | Closes the dialog                         |

**Status Messages:**
- IDS_12 (form-specific): `Performing Deactivation, Please Wait...`
- IDS_21 (form-specific): `Deactivation was successful.`

**Error Messages:**
- IDS_30 (form-specific): `The online deactivation process is not allowed. Please try manual deactivation.`
- IDS_31 (form-specific): `The deactivation process is not allowed. Cannot continue`
- IDS_13 (form-specific): `An error occurred while reading server response:`
- IDS_14 (form-specific): `An error occurred while performing deactivation. Try Manual Deactivation instead.`
- IDS_16 (form-specific): `Invalid server response format`

---

### 3.6 EULA Dialog (`TFEULA`)

**Trigger:** First run, presented before activation key entry.  
**Window Style:** `bsDialog`  
**Size:** 3,895 bytes

**Content:** Standard End User License Agreement text for SchedPro Tool (content stored in the `[Multilines]` section of the Zstd configuration as HTML, with `%APPTITLE%` placeholder substituted at runtime).

**Buttons:**
| Button Label | Action                          |
|--------------|---------------------------------|
| Accept       | Proceeds to activation          |
| Decline      | Exits the application           |

---

### 3.7 TMS Login Form (`TTMSLOGINFORM`)

**Trigger:** Conditional — if the application uses TMS Software's cloud or subscription services.  
**Size:** 21,162 bytes (large, complex form)  
**Component:** TMS Software third-party Delphi UI library

**Purpose:** Authentication for TMS-based user account (may be used for optional cloud features, update subscriptions, or additional license management).

---

### 3.8 WebUpdate Wizard (`TWUWIZ`)

**Trigger:** On startup (if auto-update enabled) or manual "Check for Updates" action.  
**Window Style:** `bsDialog`  
**Size:** 74,055 bytes (largest UI form)  
**Font:** Verdana  
**Date format:** `dd/mm/yyyy` (European)

**Components:**
- `TShape` (Shape1) — decorative background shape with button-face color
- `TImage` (Billboard) — logo/branding bitmap image (full-color TBitmap embedded in DFM)
- Multi-page wizard with progress stages

**WebUpdate Settings (from TDATAMODULE1):**
- `WebUpdate1.Agent` — HTTP user-agent string
- `WebUpdate1.DateFormat = dd/mm/yyyy`
- `WebUpdate1.TempDirectory` — temporary download location
- `WebUpdate1.Version` — current application version string

**Wizard Steps:**
1. Welcome / checking for updates
2. Downloading update (with progress bar)
3. Installing update
4. Completion / restart required

---

### 3.9 Data Module (`TDATAMODULE1`)

**Type:** Non-visual `TDataModule` (Delphi application data container)  
**Size:** 35,979 bytes  
**Events:** `OnCreate → DataModuleCreate`, `OnDestroy → DataModuleDestroy`

**Subcomponents:**
- `siLangDispatcher1` — language dispatcher (multi-language support)
- `siLang_DataModule1` — language string table component (version 7.9.7.2)
- `WebUpdate1` — update check/download component

**Configuration Properties:**
- `UseDefaultLanguage = True` — defaults to English
- `StoreAsUTF8 = True` — all strings stored in UTF-8
- `UseTaskMsgDlg = False` — uses standard dialogs, not Windows Task Dialog
- `LangDispatcher = siLangDispatcher1`
- `LangDelim = 1` — language field delimiter setting
- `LangNames.Strings = ["English"]` — only English language pack embedded

**String Table Summary:**
- Separator in string files: `~!@#$` (the `Delimiter` value from `[OPTIONS]`)
- IsUTF8File = 1
- 138 unique IDS_* entries across global + form-specific scopes
- Covers: workbook management, activation, dongle errors, trial notices, system requirement messages, add-in configuration, process errors, integrity check errors

---

## 4. Application Startup Flow (UI Perspective)

```
Launch SchedPro_Tool64.exe
        │
        ▼
[System Checks]
  ├── Excel installation check (IDS_27, IDS_276-278)
  ├── Excel version compatibility (IDS_17, IDS_33, IDS_35, IDS_36)
  ├── Not running as admin (IDS_79)
  └── Single instance check (IDS_81)
        │
        ▼
[License Check]
  ├── Trial? → Show trial notice (IDS_16, IDS_32)
  ├── Dongle? → Verify dongle (IDS_0-IDS_115)
  ├── Activated? → Proceed
  └── Not activated? → Show TFEULA → Show TFACTKEY
        │
        ▼
[TFWELC — Welcome Screen]
  ├── "Original Workbook" → Load fresh workbook
  ├── "Load A Recent Save" → Show save history list
  └── "Choose Save File" → File open dialog (*.xlsc)
        │
        ▼
[Excel Launch & Workbook Load]
  "Loading workbook, please wait..." (IDS_21)
  ├── Decrypt workbook in memory
  ├── Launch Excel via BoxedApp virtual process
  ├── Inject XLS Padlock add-in
  └── Open workbook in Excel
        │
        ▼
[SchedPro Scheduling Application — Excel UI Layer]
  (Inaccessible without activation key)
        │
        ▼
[Workbook Close — Save Dialog]
  "Do you want to keep changes made to the Excel workbook?" (IDS_5)
  ├── "Save Changes" (IDS_3) → Save to current .xlsc
  ├── "Save Changes As..." (IDS_4) → Save to new .xlsc
  └── "Ignore Changes" (IDS_2) → Discard modifications
```

---

## 5. File Dialogs

### Save File Dialog
- Filter: `Secure Excel Files|*.xlsc` (IDS_53)
- Caption: `Save Changes As...` (IDS_4)
- Trigger: User chooses to save scheduling data

### Load File Dialog
- Filter: `Secure Excel Files|*.xlsc` (IDS_53)
- Caption: `Choose Save...` (IDS_11)
- Trigger: User chooses to load an existing save

### Excel Path Selection Dialog
- Caption: `Choose Excel.exe file path` (IDS_281)
- Filter: `*.exe`
- Trigger: Excel not found at standard path (IDS_276-278)

---

## 6. Error Dialog Taxonomy

| Category              | IDS Keys         | Description                                     |
|-----------------------|------------------|-------------------------------------------------|
| System Requirements   | 17, 27, 33-36, 276-278 | Excel version, architecture, installation   |
| License/Activation    | 23-24, 37-38, 42, 76-77 | Key validation, expiry, conflicts          |
| Dongle                | 0-1, 26, 66-67, 83, 88-115 | USB security dongle errors               |
| Trial                 | 16, 19, 29-32    | Trial version notices                           |
| Excel Process         | 40, 60, 71-75, 81-82 | Excel launch and attachment errors          |
| Add-ins               | 44, 48-52        | Excel add-in configuration errors              |
| Save/Restore          | 41, 56-57, 62-63, 283-285 | Workbook save/restore errors             |
| Integrity             | 284, 285         | Save file corruption or tampering detection     |
| Access Control        | 18, 34, 79, 80   | Copy/print disabled, admin mode rejected        |

---

## 7. Accessibility and Internationalization

- **Language:** English (single language pack embedded)
- **Character Set:** UTF-8 throughout (`IsUTF8File=1`, `StoreAsUTF8=True`)
- **Font:** Tahoma (welcome), Verdana (WebUpdate wizard), DEFAULT_CHARSET for other forms
- **Date Format:** `dd/mm/yyyy` (European format used in WebUpdate component)
- **Localization Framework:** SiComponents siLang v7.9.7.2 (Delphi multilingual component)
- **DPI:** PixelsPerInch=96 (standard DPI, no explicit high-DPI manifest found)

---

## 8. Confidence Assessment

| UI Element                     | Confidence | Basis                                         |
|-------------------------------|------------|-----------------------------------------------|
| Form class names               | High       | Direct binary DFM resource extraction         |
| Form window styles/captions    | High       | Direct binary DFM property parsing            |
| Button labels                  | Medium     | Inferred from IDS_* string matching           |
| Form layout/positions          | Low        | DFM binary partially decoded                  |
| Excel workbook UI              | Very Low   | Behind AES-256 encryption; not accessible     |
| Startup flow sequence          | High       | Inferred from ordered error message IDs       |
| File dialog filters            | High       | Direct from IDS_53 string                     |
| Language/charset settings      | High       | Direct from TDATAMODULE1 DFM properties       |
