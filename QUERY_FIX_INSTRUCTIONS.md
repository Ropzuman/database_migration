# How to Fix the DB2 Empty Query Issue

## Problem Identified ✅

**Root Cause:** The SQL query for DB2 returns ZERO rows, leaving DB2 empty after "Get Data".

**Current Query (in Main sheet):**
```sql
SELECT * FROM _qryForExcel WHERE DocName3 like 'Kytkentalista'
```

**Why it fails:**
- The LIKE clause has no wildcards
- It's looking for EXACT match: `DocName3 = 'Kytkentalista'`
- The actual value in the database might be different:
  - Could be: 'Kytkentalista_something'
  - Could be: 'Some_Kytkentalista'
  - Could be spelled differently: 'Kytkentälista' (with ä)
  - Might have extra spaces

## Solution Options

### Option 1: Add Wildcards (RECOMMENDED)

Change the query in Main sheet to:
```sql
SELECT * FROM _qryForExcel WHERE DocName3 like '%Kytkentalista%'
```

The `%` means "match anything before or after".

### Option 2: Use Correct Spelling

If the database has 'Kytkentälista' with **ä** instead of **a**:
```sql
SELECT * FROM _qryForExcel WHERE DocName3 like '%Kytkentälista%'
```

### Option 3: Remove WHERE Clause (For Testing)

To see ALL documents (for testing only):
```sql
SELECT * FROM _qryForExcel
```

This will show you what's actually in the database.

## Step-by-Step Fix

### Step 1: Test What's in the Database

1. **Temporarily change** the DB2 query in Main sheet to:
   ```sql
   SELECT * FROM _qryForExcel
   ```

2. Click "Get Data"

3. Look at DB2 sheet - you should see ALL documents

4. Find the `DocName3` column

5. Look for entries that contain "Kytkentalista" or similar

6. **Note the exact spelling and format**

### Step 2: Update the Query

Once you know the correct value, update the query with wildcards:

```sql
SELECT * FROM _qryForExcel WHERE DocName3 like '%[correct_value]%'
```

Replace `[correct_value]` with what you found in Step 1.

### Step 3: Test

1. Click "Get Data" again
2. Check DB2 - should now have data in row 2
3. Click "Run Check"
4. Info sheet should populate!

## Quick Test

I've updated `HaeData` to show a warning message when DB2 returns no data. 

**Next time you click "Get Data":**
- If DB2 query returns no rows, you'll see a warning message
- The message will suggest adding wildcards
- Check the Immediate Window (Ctrl+G in VBA) for more details

## Common Query Patterns

```sql
-- Exact match (rarely what you want)
WHERE DocName3 = 'Kytkentalista'

-- Contains (usually what you want)
WHERE DocName3 LIKE '%Kytkentalista%'

-- Starts with
WHERE DocName3 LIKE 'Kytkentalista%'

-- Ends with
WHERE DocName3 LIKE '%Kytkentalista'

-- Case insensitive (Access is case-insensitive by default)
WHERE UCASE(DocName3) LIKE '%KYTKENTALISTA%'

-- Multiple conditions
WHERE DocName3 LIKE '%Kytkentalista%' OR DocName3 LIKE '%Kytkentälista%'
```

## After Fixing

Once DB2 has data, the Info and Revisions sheets should populate correctly with all the fixes we've already applied.
