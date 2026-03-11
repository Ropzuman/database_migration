# Pikatestausopas — 64-bit v2.0 / Quick Testing Guide

> **Luokitus / Classification:** `[ACTIVE]` — Kehittäjän testausopas
> **Versio / Version:** 2.0 — 64-bit M365
> **Kohderyhmä / Audience:** Kehittäjät / Developers

## 🧪 Fast Testing Checklist

### Pre-Test Setup ✅

1. ✅ Koodi kääntyy — kaikki moduulit ilman virheitä / Code compiles — all modules without errors
2. ✅ Käytössä on versio 2.0 (64-bit M365 -yhteensopiva) / Using version 2.0 (64-bit M365 compatible)
3. ✅ Varmuuskopio olemassa (`_archive/README_pre-v2.md` dokumentoi edellisen tilan)

---

## Test Sequence

### Test 1: Basic Workflow (5 minutes)

**Purpose:** Verify core functionality works

1. **Open workbook in Excel**
2. **Run HaeData:**
   - Press `Alt+F8` → Select `HaeData` → Run
   - ✅ Should complete with "Data brought successfully!"
   - Check DB1 and DB2 sheets populated

3. **Run Checkout:**
   - Press `Alt+F8` → Select `Checkout` → Run
   - ✅ Should complete with "Check OK!"
   - Check Info sheet has doc properties filled

4. **Run GenPrintout:**
   - Press `Alt+F8` → Select `GenPrintout` → Run
   - 👀 **Watch status bar** (bottom of Excel) for progress messages
   - ✅ Should complete with file save dialog
   - Open saved file and verify:
     - Info sheet correct
     - POSheet has data
     - Legend sheet present
     - Revisions sheet populated
     - Footers present on first 3 sheets
     - LINKING hidden or deleted (depending on checkbox)

**Expected:**

- ✅ No errors
- ✅ Faster execution than before
- ✅ Much less screen flashing
- ✅ Status bar shows progress

**If errors occur:**

- Note the error message (it should be detailed now)
- Check which step failed
- Report back with error text

---

### Test 2: Empty Dataset (2 minutes)

**Purpose:** Handle edge case gracefully

1. Modify SQL query in Main sheet to return 0 results
   (e.g., add `WHERE 1=0` to query)
2. Run `HaeData`
3. Run `Checkout`
4. Run `GenPrintout`

**Expected:**

- ✅ Should complete without "subscript out of range" errors
- ✅ Generate workbook with headers but no data rows

---

### Test 3: Hide LINKING Toggle (3 minutes)

**Purpose:** Verify both options work

1. **With checkbox CHECKED:**
   - Run GenPrintout
   - ✅ LINKING sheet should be hidden (not deleted)
   - Right-click sheet tabs → Unhide → Should see LINKING

2. **With checkbox UNCHECKED:**
   - Run GenPrintout
   - ✅ LINKING sheet should be deleted completely
   - Right-click sheet tabs → Unhide → Should NOT see LINKING

---

### Test 4: Large Dataset (Optional, 5 minutes)

**Purpose:** Verify performance improvement

1. Use query with 1000+ rows
2. **Time the old version** (if you have it)
3. **Time the new version:**
   - Run GenPrintout
   - Note completion time
   - Compare with old time

**Expected:**

- ✅ 25-30% faster than old version
- ✅ Status bar updates smoothly
- ✅ Minimal screen flashing

---

## 🔍 What to Watch For

### Good Signs ✅

- Fast execution
- Smooth progress (status bar updating)
- No screen flashing/flickering
- Clear error messages (if errors occur)
- Workbook opens immediately after generation

### Warning Signs ⚠️

- "Subscript out of range" errors → Report immediately
- "Object variable not set" errors → Report immediately
- Macro hangs/freezes → Press `Ctrl+Break`, report
- Missing sheets in output → Report which sheets missing

---

## 🐛 If You Find Bugs

**Report:**

1. Which test failed (Test 1, 2, 3, or 4)
2. Exact error message (copy full text)
3. Which step in the test
4. Dataset size (small/medium/large)

**Quick recovery:**

```powershell
# Return to pre-optimization code
git checkout comments
```

Then tell me what went wrong and I'll fix it!

---

## 📊 Performance Comparison Template

Use this to track improvements:

| Metric | Old Version | New Version | Improvement |
|--------|-------------|-------------|-------------|
| HaeData time | ___ sec | ___ sec | ___% |
| Checkout time | ___ sec | ___ sec | ___% |
| GenPrintout (100 rows) | ___ sec | ___ sec | ___% |
| GenPrintout (1000 rows) | ___ sec | ___ sec | ___% |
| Screen flashing | Lots | Minimal | ~90% ↓ |
| User experience | OK | Better | ✅ |

---

## ✅ Success Criteria

**Optimization is successful if:**

- [x] All 3 main macros run without errors
- [x] Generated printouts are identical to old version (content-wise)
- [x] Execution is noticeably faster
- [x] Much less screen flashing
- [x] Status bar shows progress
- [x] Error messages are clear (if errors occur)

---

## 🎯 Expected Results Summary

**Before optimization:**

- GenPrintout: 15-20 sec (1000 rows)
- Lots of screen flashing
- No progress feedback

**After optimization:**

- GenPrintout: **10-14 sec (1000 rows)** ⚡
- Minimal flashing 🎨
- Status bar progress 📊

---

**Ready to test!** Start with Test 1 and let me know how it goes! 🚀
