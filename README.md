# Schichtbuch → SAP Sync (VBScript)

Automate the double work of logging maintenance tasks twice: once in the Excel **Schichtbuch** and once in **SAP PM**.  
This VBScript scans the shift logbook Excel, identifies entries not yet processed in SAP, and **confirms, completes, or cancels** the corresponding SAP work orders (IW41 / IW32), either fully automatically or with user confirmations.

> 🧑‍🏭 Typical users: Maintenance & Reliability (M&R) planners, supervisors, managers

> ⚙️ Runs on: Windows, SAP GUI for Windows (with scripting), Microsoft Excel

---

## Table of contents

- [Features](#features)
- [How it works](#how-it-works)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Excel layout & data mapping](#excel-layout--data-mapping)
- [Command-line usage](#command-line-usage)
- [Parameter defaults](#parameter-defaults)
- [Run examples](#run-examples)
- [Logging](#logging)
- [Safety checks & guardrails](#safety-checks--guardrails)
- [Known limitations](#known-limitations)
- [Troubleshooting](#troubleshooting)

---

## Features

- 🔍 **Reads the Schichtbuch Excel** in bulk (fast, minimal COM calls)
- 👷 **Maps employee names → SAP personnel numbers** using sheet 2 (lenient name matching)
- 🕒 **Converts Excel time fractions** into proper timestamps and correctly handles **overnight work**
- 🌍 Converts all timestamps to **UTC** using Windows ActiveTimeBias
- 🗃️ **Confirms work orders** in IW41 including short texts, work duration, start/end, and final confirmation settings
- ❌ **Cancels orders** in IW32 when the Excel log marks them as cancelled
- 🧹 **Optionally completes (TECO)** confirmed work orders
- 📄 **Writes back results** to the Excel (column 16) only if no existing message is present
- 🔁 **Optional helper:** copies upcoming PMs from **IW38** into the Excel (layout-dependent)
- 🛡️ Robust SAP wrappers (`SafeFindById`, `SafeSetText`, `SafeSendVKey`, etc.)  
- 🪵 **Detailed log files** in `./logs/`

---

## How it works

1. Script initializes logging, reads runtime arguments, loads timezone offset.
2. Determines which Excel file to open based on:
   - `filePath` argument
   - OR `useCurrentExcel=yes` (builds SharePoint path)
   - OR file-open dialog (default)
3. Excel is opened; sheet 1 & sheet 2 are bulk-read for performance.
4. Each row in sheet 1 is validated and processed:
   - Checks WO, employee, status, times
   - Cancels orders in IW32 when needed
   - Determines SAP status and skip-conditions
   - Confirms WOs via IW41 for one or multiple employees
   - Performs TECO if needed
5. Writes result messages to column 16
6. Logs everything and exits cleanly

---

## Prerequisites

- Windows with **Windows Script Host (WSH)**
- SAP GUI 7.x+ with **scripting enabled** (client + server)
- Microsoft Excel installed
- SAP access to IW32, IW33, IW38, IW41
- IW38 helper requires ALV layout with **technical field names**

---

## Installation

1. Place these files in the same directory:
   - [`Schichtbuch script.vbs`](https://github.com/TXAE/CGL/blob/main/Schichtbuch%20script.vbs)
   - [`SAP Login.vbs`](https://github.com/TXAE/CGL/blob/main/SAP%20Login.vbs)
2. Ensure folder is writable (script creates `./logs/`)
3. Ensure SAP GUI scripting is enabled

---

## Excel layout & data mapping

### Sheet 1 (Schichtbuch)

| Column | Description |
|--------|-------------|
| 1 | Date (Tag) |
| 3 | WO_Nr (must be 9 digits starting with 4) |
| 5 | Employee(s) separated by `/` |
| 7 | Bemerkung |
| 8 | Fehlerbeschreibung |
| 9 | Massnahme (max 40 chars) |
| 10 | Startzeit (Excel fraction) |
| 11 | Endzeit (Excel fraction) |
| 12 | DauerInH (fraction or time) |
| 15 | Status (matches sheet 2) |
| 16 | Script output message |

### Sheet 2 (Mappings)

- Column A: Employee name
- Column B: Personnel number
- Cell E4: "done" text
- Cell E5: "cancelled" text

---

## Command-line usage

```
cscript //nologo "Schichtbuch script.vbs" [filePath=<path_or_url>] [autoConfirm=yes|no] [useCurrentExcel=yes|no]
```

---

## Parameter defaults

### `filePath`
**Default:** not provided  
**Behavior:**
- If `useCurrentExcel=yes` → script builds current-month SharePoint path
- Otherwise → shows Excel file-open dialog

### `autoConfirm`
**Default:** not provided  
**Behavior:** script asks user:
- **Yes** → automatic confirmations
- **No** → ask before each confirmation

### `useCurrentExcel`
**Default:** no  
**Behavior:**
- `yes` → build SharePoint path for current month's Schichtbuch
- `no` → normal file selection process

---

## Run examples

### Interactive
```
cscript //nologo "Schichtbuch script.vbs"
```

### Fully automatic
```
cscript //nologo "Schichtbuch script.vbs" filePath="C:\Data\Schichtbuch.xlsx" autoConfirm=yes
```

### Auto-select current month
```
cscript //nologo "Schichtbuch script.vbs" useCurrentExcel=yes autoConfirm=yes
```

---

## Logging

Stored in `./logs/<script>_<user>_<timestamp>.log`.

---

## Safety checks & guardrails

- Hard skip conditions: missing WO, wrong WO format, message already present, status missing
- Soft skips: missing employee, missing times, invalid conversions
- SAP skip conditions: purchases found, multiple operations
- SAP wrappers abort cleanly on layout or scripting errors

---

## Known limitations

- SAP GUI layout differences require adjustments
- IW38 helper depends on technical ID layout
- Massnahme truncated to 40 chars
- Buffered writes disabled for reliability

---

## Troubleshooting

### SAP object not found
Check SAP GUI scripting and correct transaction screen.

### Employee not found
Add employee + personnel number to sheet 2.

### Time wrong
Ensure Excel uses real time values or fractions.
