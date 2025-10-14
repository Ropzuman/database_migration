# Debug Info Sheet Issue

## What We Added

Added comprehensive debug output to see exactly what's happening:

### HaeDocTiedot Debug
Shows what values are read from DB2 into the DI variables.

### VaihdaInfo Debug
Shows each comment found and what value it's being set to.

## How to Test

### Option 1: Run Quick Test
1. VBA Editor (Alt+F11)
2. Clear Immediate Window (Ctrl+G, right-click → Clear All)
3. Press F5
4. Select `QuickInfoSheetTest`
5. Run

### Option 2: Normal Workflow
1. Clear Immediate Window (Ctrl+G in VBA, right-click → Clear All)
2. Click "Get Data"
3. Click "Run Check"
4. Check Immediate Window

## What to Look For

The Immediate Window will show:

```
HaeDocTiedot: Populated variables:
  DICustomer: 'Fortum'
  DIMill: 'Nuijalan lämpölaitos'
  DIProject: 'Fortum Nuijalan lämpölaitos'
  ...

VaihdaInfo: Comment 'customer' at $C$5
  -> Set to: Fortum
VaihdaInfo: Comment 'mill' at $C$6
  -> Set to: Nuijalan lämpölaitos
...
```

## What We're Looking For

1. **Are the DI variables correct?** (Should be Fortum for Customer, Nuijalan lämpölaitos for Mill)
2. **What comments does the Info sheet actually have?** 
3. **Which comment is matching to which cell?**

## Possible Issues

### Issue 1: Comment Text is Wrong
If the comment says something other than "customer" (like "Customer" with capital C, or "&&Customer"), we need to handle that.

### Issue 2: Multiple Cells Have Same Comment
If multiple cells have the "customer" comment, they'll all get the same value.

### Issue 3: Wrong Variable Being Used
The mapping might be backwards somewhere.

## Next Steps

Run the test and copy the **complete Immediate Window output** here. That will tell us:
- What values are in the variables
- What comments exist on the Info sheet
- Where each value is being written

Then we can fix the exact issue!
