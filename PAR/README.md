# PAR

Project Asset Register (PAR) is an Excel/VBA workbook codebase for compiling project asset data, producing cost-of-project details, and generating Assetic export worksheets for new, renewed, and disposed assets.

## Overview
- Populates a handover cost report by scanning asset sheets, building cost lines, and appending write-off items.
- Generates Assetic extract worksheets for new assets, renewals (CapEx), and disposals, with headers and row-level mappings.
- Provides navigation macros to jump between summary and data sheets, and utilities to hide/unhide lookup tables.
- Exports the active worksheet to CSV in `C:\Assetic_Extract`.

## Modules

### Module1.bas
`Populate_Handover_Cost` builds the "Cost of Project Details" section on the handover cost sheet:
- Locates the target insertion row by finding "Cost of Project Details" and the "Category" row.
- Clears a range of existing rows before repopulating.
- Reads column positions from the New Assets and Renewed Assets sheets based on header labels (e.g., Asset Class/Type/ID, Quantity, Unit/Total Cost, Allocate Project, Capitalise This, Valuation Record ID, Useful Life, Asset SubClass/SubType, Component Name).
- Populates cost lines for each qualifying new/renewed asset, inserting formulas for allocations, overheads, and capitalization.
- Appends write-off rows by scanning the Project Wide Costs sheet.
- Uses blank-row counters to stop after a run of empty lines and guards against missing template columns.

### Module2.bas
Sheet navigation and visibility helpers:
- `View_*` macros select summary, construction, project-wide, handover, new, renewed, disposed, and transactions sheets.
- `Hide_Lookup` and `UnHide_Lookup` toggle visibility for reference/lookup sheets (asset class/type, CoA schema, treatment types, UoM, and other lookups).

### Module3.bas
`Populate_AsseticNewAssets` generates Assetic extract sheets for new assets:
- Clears existing rows in Assetic output sheets.
- Reads project code/description from named ranges.
- Detects required input columns on the New Assets sheet.
- Writes header rows for Assetic New Assets, Components, Network Measures, and Valuations sheets.
- Maps each eligible asset row into all four Assetic sheets, including valuation patterns for specific subclasses and unit-rate calculations.
- Renames output sheets with the project code prefix.

### Module4.bas
`Populate_AsseticRenewedAssets` produces the Assetic CapEx Renewals worksheet:
- Clears existing output rows and reads project code/description.
- Detects input columns on the Renewed Assets sheet, including renewal percentages, condition rating, and treatment type.
- Writes the CapEx Renewals header and maps each eligible renewal row to Assetic fields.
- Sets valuation patterns based on subclass and derives treatment naming.
- Renames the output sheet with the project code prefix.

### Module5.bas
`Populate_AsseticDisposedAssets` generates Assetic Disposed Assets and Disposed Valuations:
- Clears existing output rows and reads project code/description.
- Detects required columns on the Disposed Assets sheet (IDs, components, valuation IDs/dates/types, disposal details, comments).
- Writes headers and maps disposal rows (full asset disposals only) into both output sheets.
- Renames output sheets with the project code prefix.

### Module6.bas
`Generate_Cost_and_Assetic` is the orchestration macro:
- Calls the handover cost generator.
- Temporarily shows all Assetic output sheets.
- Runs new/renewed/disposed Assetic population routines.
- Re-hides Assetic sheets and restores screen updates.

### Module7.bas
`CurSheet_To_CSV` exports the active worksheet to CSV:
- Ensures `C:\Assetic_Extract` exists.
- Copies the active sheet to a new workbook and saves it as CSV.
- Displays the output location.
