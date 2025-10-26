# AcadDATA – Late Binding Log and Constant Map

Date: 2025-10-26

## TL;DR
- We use late binding (Object) to avoid AutoCAD version references; this makes the workbook portable across machines/versions.
- With late binding, VBA doesn’t know AutoCAD’s enum names, so we define the numeric values ourselves.
- Selection mode constant must be correct: `acSelectionSetAll = 5` (not 2). Value 2 is Fence selection and requires points, which caused “Invalid argument Mode in Select”.
- Filter type array must be `Integer()`, not `Long()`. Using `Long` throws “Invalid argument FilterType in Select”.

## Why constants are hard-coded
When you don’t set a reference to the “AutoCAD x.x Type Library,” VBA can’t resolve symbols like `acSelectionSetAll`, `acModelSpace`, `acMax`, or `ac2013_dwg`. AutoCAD’s COM API still expects those enum values; we provide them numerically so the calls succeed while keeping readable names in code.

Late binding benefits for this tool:
- Version independence across AutoCAD installs (no “Missing: AutoCAD xx Type Library”).
- Simple deployment to multiple machines.
Trade-offs:
- No IntelliSense and no compile-time enum checking.
- We must define numeric constants and be precise.

## Constant map used by this project
These values mirror AutoCAD’s type library (verified via early-binding checks and runtime behavior).

### Selection modes (AcSelect)
- `acSelectionSetPrevious = 4` — use previous selection
- `acSelectionSetAll = 5` — select all (programmatic)
- Note: `2` is Fence selection and requires fence points. Passing `Empty` for points with 2 results in: “Invalid argument Mode in Select”.

### Active space (AcActiveSpace)
- Project value: `acModelSpace = 1` (used to force model space before zoom)
- If you enable early binding, confirm your enum values in the Immediate window (see “Verification”). Some builds report `acModelSpace = 0`. Use the value your AutoCAD type library reports.

### Window state (AcWindowState)
- `acMax = 3` — maximize window

### Zoom (AcZoomScaleType)
- `acZoomScaledRelative = 3` — relative zoom

### SaveAs types (AcSaveAsType)
- `ac2004_dwg = 24`
- `ac2007_dwg = 36`
- `ac2010_dwg = 48`
- `ac2013_dwg = 60`
- `acNative = 60` — in current projects this maps to the default/native SaveAs version supported by the installed AutoCAD

## Known runtime errors and fixes
- Error: “Invalid argument FilterType in Select”
  - Cause: `FilterType()` declared As Long.
  - Fix: Declare `Dim FilterType() As Integer` and pass it with `FilterData()` to Select/SelectOnScreen.

- Error: “Property let procedure not defined and property get procedure did not return an object”
  - Cause: Using late binding without defining AutoCAD constants (enums evaluated to Empty/Variant).
  - Fix: Define required constants numerically in the module.

- Error: “Invalid argument Mode in Select”
  - Cause: `acSelectionSetAll` incorrectly set to 2 (Fence) while passing Empty for Point1/Point2.
  - Fix: Set `acSelectionSetAll = 5`.

## Selection patterns with late binding
- Use programmatic selection (all, with filter):
  - `Joukko.Select acSelectionSetAll, Empty, Empty, FilterType, FilterData`
- Use user selection on screen (respecting filter):
  - `Joukko.SelectOnScreen FilterType, FilterData`
- For previous selection reuse (non-interactive):
  - `Joukko.Select acSelectionSetPrevious, Empty, Empty, FilterType, FilterData`

## Verification: how to check enum values safely
If you want to verify exact numbers on your machine:
1. In the VBA editor, Tools → References → check your installed “AutoCAD x.x Type Library”.
2. In the Immediate window, query values (examples):
   - `? acSelectionSetAll` → expected 5
   - `? acSelectionSetPrevious` → expected 4
   - `? acMax` → expected 3
   - `? ac2013_dwg` → expected 60
   - `? acModelSpace` → confirm value reported by your AutoCAD build (0 or 1; use that in constants if needed)
3. Remove the reference again to keep the workbook version-independent.

## Late binding vs early binding for this workbook
- Keep late binding if users have different AutoCAD versions or you distribute widely.
- Consider early binding for a standardized environment to regain IntelliSense and remove manual constants. If you switch:
  - Add the AutoCAD reference.
  - Remove/ignore hard-coded constants (the enum names will resolve).
  - Compile and test TuoDATA/VieDATA.

## References
- AutoCAD ActiveX (ObjectARX/OARX) Reference — SelectionSet.Select method and enumerations
  - AutoCAD help portal: https://help.autodesk.com/view/OARX/ (navigate to ActiveX Reference → Methods/Enums)
- VBA Type Conversion rules (why Integer vs Long matters with COM arrays):
  - https://learn.microsoft.com/office/vba/language/reference/user-interface-help/type-conversion-functions

## Change log (late-binding related)
- 2025-10-26
  - Fixed `acSelectionSetAll` to 5 (was 2; caused Mode error).
  - Documented that `FilterType()` must remain Integer.
  - Clarified Select vs SelectOnScreen usage with filters under late binding.
