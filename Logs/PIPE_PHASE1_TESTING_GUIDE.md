# PIPE Database - Phase 1 Testing Guide

**Date**: November 12, 2025  
**Status**: Ready for Testing  
**Commits**: 9dfb7a5 (initial), 5641463 (optimization)

---

## What's Been Completed

### Files Migrated and Optimized (5 of 11)

1. **Koodit.bas** - Core module with AutoCAD integration
2. **Form_Linkkien vaihto.cls** - Link updater utility
3. **Form_zFunc.cls** - zDetails table management utilities
4. **Form_USysFlowPickNo.cls** - Flow block pipeline number picker
5. **Form_frmOpenPIPELINE.cls** - Pipeline segment selector

### Changes Applied

- ✅ VBA7/64-bit compatibility (API declarations)
- ✅ Explicit DAO typing (all database objects)
- ✅ Comprehensive error handling (15+ procedures)
- ✅ Full documentation (module, form, function levels)
- ✅ Finnish encoding fixes (8 instances)
- ✅ Performance optimizations (12+ improvements)
- ✅ Transaction support (batch operations)
- ✅ Bug fixes (navigation, ListIndex, parameter validation)

---

## Testing Instructions

### Prerequisites

- PIPE database opened in Access
- **Optional**: AutoCAD running for full AutoCAD integration tests
- VBA Editor access (Alt+F11)

---

## Test 1: Compilation Test (CRITICAL)

**Purpose**: Verify all code compiles without errors

### Steps

1. Open PIPE database in Access
2. Press `Alt+F11` to open VBA Editor
3. Go to **Debug** → **Compile VBA Project**
4. Check for any compilation errors

### Expected Result

- ✅ No compilation errors
- ✅ Status bar shows "Compilation completed successfully"

### If Errors Occur

- Note the exact error message and line number
- Report back before continuing

---

## Test 2: Koodit.bas Module

### Test 2a: SetStartup Function

**Purpose**: Verify user login tracking

### Steps

1. In VBA Editor, open `Koodit` module
2. In Immediate Window (Ctrl+G), type: `SetStartup`
3. Press Enter
4. Open `UsysUsers` table
5. Check for new entry with current timestamp

### Expected Result

- ✅ New record in UsysUsers with:
  - Network username
  - Database username
  - Computer name
  - Current timestamp

---

### Test 2b: POIMI Function

**Purpose**: Verify string parsing utility

### Steps

In Immediate Window, test:

```vba
? POIMI("AREA-123-VALVE", 1)
? POIMI("AREA-123-VALVE", 2)
? POIMI("AREA-123-VALVE", 3)
```

### Expected Results

```
AREA
123
VALVE
```

---

### Test 2c: AvaaBlock (MANUALVALVES) - Requires AutoCAD

**Purpose**: Verify AutoCAD block opening from valve table

### Steps

1. Open `MANUALVALVES` table in datasheet view
2. Select a record that has drawing information
3. Click **Add-Ins** → **Etsi kohde** (or call AvaaBlock)

### Expected Result

- ✅ AutoCAD opens (or activates if already open)
- ✅ Drawing opens
- ✅ Zooms to valve block
- ✅ Block is highlighted

### If AutoCAD Not Available

- Test passes if error message: "Käynnissä olevaa AutoCADiä ei löytynyt!"

---

### Test 2d: AvaaBlock (PIPELINES) - Requires AutoCAD

**Purpose**: Verify pipeline segment selector

### Steps

1. Open `PIPELINES` table in datasheet view
2. Select a pipeline record
3. Click **Add-Ins** → **Etsi kohde**

### Expected Result

- ✅ `frmOpenPIPELINE` form opens
- ✅ Shows list of pipeline segments
- ✅ If only one segment: automatically opens drawing
- ✅ If multiple segments: allows selection

---

## Test 3: Form_Linkkien vaihto

**Purpose**: Verify link updater utility

### Steps

1. Open form `Linkkien vaihto`
2. Click the update button
3. Observe the result message

### Expected Results

**Scenario A** (No linked tables):

- ✅ Message: "Ei linkitettyjä tauluja löytynyt"

**Scenario B** (Linked tables with different path):

- ✅ Message: "Päivitetty X linkkiä" (where X is number of relinked tables)
- ✅ All linked tables now point to current directory

**Scenario C** (Linked tables already in current path):

- ✅ Message: "Ei linkkejä päivitettävänä"

---

## Test 4: Form_zFunc Utilities

### Test 4a: Block Placement Numbering

**Purpose**: Verify zDetails block placement assignment

### Steps

1. Open form `zFunc`
2. Click **Command3** button (block placement)
3. Check confirmation message

### Expected Result

- ✅ Message: "Homma hoidettu"
- ✅ zDetails records have sequential block placement numbers (14, 15, 16...)
- ✅ Numbers reset for each new loop

---

### Test 4b: Empty String to Null Conversion

**Purpose**: Verify batch null conversion with transaction support

### Steps

1. Open form `zFunc`
2. Select a table from `KaikkiTaulukot` dropdown
3. Click **Command11** button
4. Observe progress message

### Expected Result

- ✅ Message: "Valmis: X riviä käyty, Y päivitetty"
- ✅ All empty strings in table converted to Null
- ✅ No errors during processing

### Performance Note

- Should be faster than previous version due to:
  - Transaction support
  - Single Edit/Update per record (not per field)

---

### Test 4c: Orphan Record Deletion

**Purpose**: Verify orphaned zDetails cleanup with transaction support

### Steps

1. Open form `zFunc`
2. Click **Command5** button
3. Observe hourglass cursor (should appear once, not flicker)
4. Check result message

### Expected Results

**Scenario A** (No orphans):

- ✅ Message: "Ei vanhentuneita tietueita"

**Scenario B** (Orphans found):

- ✅ Message: "Poistettu X vanhentunutta tietuetta"
- ✅ Orphaned records removed from zDetails
- ✅ Hourglass appears once (not per record)

---

### Test 4d: Form Load (Table List Population)

**Purpose**: Verify table list filtering

### Steps

1. Open form `zFunc`
2. Click `KaikkiTaulukot` dropdown

### Expected Result

- ✅ List contains user tables only
- ✅ No system tables (USys*, MSys*, Dev*)
- ✅ No tables with "old" in name
- ✅ First table selected by default

---

## Test 5: Form_USysFlowPickNo - Requires AutoCAD

**Purpose**: Verify flow block pipeline number picker

### Test 5a: Form Connection

### Steps

1. Ensure AutoCAD is running with a drawing containing flow blocks
2. Open form `USysFlowPickNo`

### Expected Results

**Scenario A** (AutoCAD running with flow blocks):

- ✅ Form opens successfully
- ✅ Form caption shows: "BlockName: 1/X" (where X is total blocks)
- ✅ First flow block highlighted in AutoCAD
- ✅ AutoCAD zooms to block
- ✅ TRefNo shows current pipeline reference

**Scenario B** (AutoCAD not running):

- ✅ Error: "Käynnissä olevaa AutoCADiä ei löytynyt!"
- ✅ Form closes gracefully

**Scenario C** (No flow blocks found):

- ✅ Message: "Ei tasearvoblokkeja"
- ✅ Form closes gracefully

---

### Test 5b: Navigation (Next/Previous)

### Steps

1. With form open, click **BSeuraava** (Next) button multiple times
2. Click **BEdellinen** (Previous) button multiple times
3. Navigate to last block, click Next again

### Expected Results

- ✅ Next button cycles through blocks (1 → 2 → 3 → ... → last → 1)
- ✅ Previous button cycles backwards (last ← ... ← 3 ← 2 ← 1 ← last)
- ✅ AutoCAD highlights current block
- ✅ Form caption updates with current position
- ✅ TRefNo shows pipeline reference for each block

### Bug Fix Verification

- ✅ No error when clicking Next on last block (wraparound works correctly)

---

### Test 5c: Find Next Empty

### Steps

1. Click **BSeurT** (Seur.Tyhjä) button
2. Form should jump to next flow block with empty pipeline reference

### Expected Results

**Scenario A** (Empty blocks found):

- ✅ Skips blocks with pipeline references
- ✅ Stops at first empty block
- ✅ TRefNo is empty

**Scenario B** (No empty blocks):

- ✅ Message: "Sellaisia virtausarvoblokkeja ei löytynyt, joissa putkilinja olisi tyhjä."

---

### Test 5d: Pick Pipeline from Drawing

### Steps

1. Navigate to empty flow block
2. Click **BPoimi** button
3. AutoCAD prompts: "Poimi putkilinja..."
4. Click on a pipeline block in AutoCAD drawing

### Expected Results

- ✅ AutoCAD window activates
- ✅ Pick prompt appears
- ✅ After clicking pipeline block:
  - TRefNo populated with pipeline reference
  - Flow block attribute updated in AutoCAD
  - Focus returns to Access form

### Supported Block Types

- PIPELINE, PIPELINE_F, ARAPIPEL, BAHPIPEL, METSO_PIPE
- Other blocks with POS.NO, CUSTPOS, or SDPOS attributes

---

### Test 5e: Auto-Search Pipelines

### Steps

1. Click **BAutohaku** button
2. Wait for processing (iterates all blocks)
3. Check result message

### Expected Results

**Scenario A** (All pipelines found):

- ✅ Message: "Kaikki linjat on käyty läpi. Kaikkiin virtausarvoblokkeihin löytyi putkilinjanumero."
- ✅ All flow blocks have pipeline references

**Scenario B** (Some pipelines not found):

- ✅ Message: "Kaikki linjat on käyty läpi. Yksi tai useampia virtausarvoblokkeja jäi ilman putkilinjanumeroa."
- ✅ Form navigates to first empty block
- ✅ User can manually pick or enter pipeline reference

### Search Logic

- Creates 56x13 unit window from each flow block insertion point
- Searches for PIPELINE, BAHPIPEL, ARAPIPEL blocks
- Extracts DEP+FSB+LINE from first found pipeline

---

### Test 5f: Manual Entry

### Steps

1. Navigate to any flow block
2. Type pipeline reference in TRefNo textbox
3. Press Tab or click elsewhere

### Expected Result

- ✅ AutoCAD block attribute updates immediately
- ✅ No errors

---

### Test 5g: Form Close

### Steps

1. Click **BSulje** (Close) button

### Expected Result

- ✅ Form closes
- ✅ Selection set deleted from AutoCAD
- ✅ No errors

---

## Test 6: Form_frmOpenPIPELINE

**Purpose**: Verify pipeline segment selector

### Test 6a: Single Segment Pipeline - Requires AutoCAD

### Steps

1. Open `PIPELINES` table
2. Find pipeline with only one segment
3. Select record and click **Add-Ins** → **Etsi kohde**

### Expected Results

- ✅ Form opens briefly
- ✅ AutoCAD opens/activates automatically
- ✅ Drawing opens
- ✅ Zooms to pipeline block
- ✅ Form closes automatically
- ✅ No user interaction needed

---

### Test 6b: Multi-Segment Pipeline - Requires AutoCAD

### Steps

1. Open `PIPELINES` table
2. Find pipeline with multiple segments in different drawings
3. Select record and click **Add-Ins** → **Etsi kohde**

### Expected Results

- ✅ Form opens and stays open
- ✅ Lista listbox populated with segments
- ✅ Each row shows: PATH, Drawing, FROM, TO info
- ✅ First segment selected by default
- ✅ Click **Command3** (Open) button:
  - AutoCAD opens/activates
  - Selected segment drawing opens
  - Zooms to pipeline block
- ✅ Click **Command4** (Close) button: Form closes

---

### Test 6c: No Pipeline Data

### Steps

1. Manually open form: `DoCmd.OpenForm "frmOpenPIPELINE", , , , , , "XX,999"`

### Expected Result

- ✅ Error: "Putkilinjatietoja ei ole tuotu virtauskaavioista"
- ✅ Form closes

---

### Test 6d: Invalid OpenArgs

### Steps

1. Open form without OpenArgs or with malformed args

### Expected Results

**Scenario A** (No OpenArgs):

- ✅ Error: "Avaa tämä lomake putkilinjataulukon kautta..."
- ✅ Form closes

**Scenario B** (Malformed OpenArgs):

- ✅ Error: "Virheelliset parametrit"
- ✅ Form closes

### Bug Fix Verification

- ✅ ListIndex used correctly (no +1 offset error)

---

## Test 7: Finnish Character Encoding

**Purpose**: Verify all Finnish characters display correctly

### Steps

1. Open each form and module
2. Check messages and comments for:
   - ä, ö, Ä, Ö characters
   - Common words: käyty, löytynyt, tämä, sisältänyt, Käynnissä

### Expected Result

- ✅ All Finnish characters display correctly
- ✅ No � (question mark/box) characters

### Files to Check

- Form_USysFlowPickNo: "käyty läpi", "löytynyt", "sisältänyt", "Käynnissä"
- Form_frmOpenPIPELINE: "tämä lomake"
- Form_zFunc: "lisättävät"
- Koodit.bas: Error messages

---

## Performance Comparison Tests (Optional)

### Test 8a: Empty String Conversion Performance

**Purpose**: Verify transaction optimization

### Steps

1. Note a large table name
2. Time the Command11_Click operation
3. Check progress message

### What to Look For

- ✅ Progress message shows records processed vs updated
- ✅ Operation completes faster than expected
- ✅ Only updated records counted in "päivitetty" count

### Technical Notes

- Old version: Edit/Update for every field (slow)
- New version: Single Edit/Update per record (fast)
- Transaction support: Rollback on error

---

### Test 8b: Orphan Deletion Performance

**Purpose**: Verify hourglass optimization

### What to Look For

- ✅ Hourglass appears once at start
- ✅ Hourglass does NOT flicker per iteration
- ✅ Progress message shows count of deleted records

### Technical Notes

- Old version: Hourglass per iteration (annoying)
- New version: Hourglass once (clean)
- Transaction support: All-or-nothing deletion

---

## Common Issues and Solutions

### Issue 1: Compilation Error

**Symptom**: Compile error on API declarations

**Solution**: Ensure Access is 64-bit version. Check:

```vba
? Application.Version
```

Should be version 16.0 or later for Office 365/2016+

---

### Issue 2: AutoCAD Not Found

**Symptom**: "Käynnissä olevaa AutoCADiä ei löytynyt!"

**Solution**:

- Start AutoCAD before opening forms
- Open a drawing (not just AutoCAD)
- Verify GetObject works in Immediate Window:

```vba
? GetObject(, "AutoCAD.Application").Name
```

---

### Issue 3: DAO Reference Missing

**Symptom**: "User-defined type not defined" on DAO.Database

**Solution**:

1. VBA Editor → Tools → References
2. Check "Microsoft DAO 3.6 Object Library" (or newer)
3. Click OK
4. Recompile

---

### Issue 4: Block Not Found in AutoCAD

**Symptom**: "Kuvasta ei löytynyt kohdetta tietokannan tiedoilla"

**Solution**:

- Handle value in database may be outdated
- Drawing may have been regenerated
- Block may have been deleted
- Re-import data from AutoCAD

---

## Success Criteria

### Minimum Requirements (Must Pass)

- ✅ All code compiles without errors
- ✅ No runtime errors in any tested procedure
- ✅ Finnish characters display correctly
- ✅ Transaction rollback works on error

### Full Test Success (Ideal)

- ✅ All minimum requirements met
- ✅ AutoCAD integration works (if AutoCAD available)
- ✅ All navigation works correctly
- ✅ Performance improvements noticeable
- ✅ All bug fixes verified

---

## Reporting Results

### What to Report

**For Each Test**:

1. Test number and name
2. Result: ✅ PASS or ❌ FAIL
3. If FAIL: Error message, steps to reproduce
4. Screenshots (if helpful)

**Overall**:

- Environment: Access version, OS version
- AutoCAD: Available? Version?
- Any observations or suggestions

---

## After Testing

### If All Tests Pass

✅ **Proceed to Phase 2**: Migrate remaining 6 files

### If Issues Found

❌ **Report issues**: Agent will fix before continuing

---

## Quick Reference

### Files Modified

```
Access/PIPE/Koodit.bas
Access/PIPE/Form_Linkkien vaihto.cls
Access/PIPE/Form_zFunc.cls
Access/PIPE/Form_USysFlowPickNo.cls
Access/PIPE/Form_frmOpenPIPELINE.cls
PIPE_PROGRESS_PHASE1.md
```

### Key Improvements

- 12+ performance optimizations
- Transaction support (2 operations)
- 15+ error handlers
- 8 encoding fixes
- 3 bug fixes
- 1,140+ lines migrated/optimized

---

*Testing Guide Generated: November 12, 2025*  
*Ready for Phase 1 Testing*
