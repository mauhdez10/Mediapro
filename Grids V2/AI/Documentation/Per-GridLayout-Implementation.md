# Per-Grid Layout Settings Implementation - Summary

## What Was Implemented

This feature allows each grid type (CATV, TVD, PASIONES_LATAM, PASIONES_US, TODO_NOVELAS, REV_TV, SPORTYNET) to have its own customizable layout settings instead of using global hardcoded values.

## Changes Made

### 1. FormatGrids.ps1 (Main Formatter)

#### Added `Get-GridLayout()` Helper Function
- **Location:** After `Get-TabRule()` function (around line 234)
- **Purpose:** Retrieves per-grid layout settings from `$ManagedSettings.Layouts`
- **Returns:** An ordered hashtable with 5 properties:
  - `FontSize` (int)
  - `DefaultRowH` (double)
  - `HeaderRowH` (double)
  - `SmallColW` (double)
  - `DefaultColW` (double)
- **Fallback:** Returns `$DefaultLayout` if no grid-specific settings found

#### Updated Managed Settings Block
- **Location:** Lines 8-35 (managed by Settings UI)
- **Added:** `Layouts` section with default values for all 7 grid types
- **Property Names:** Must match Settings UI expectations:
  - `FontSize`
  - `DefaultRowHeight` (not `DefaultRowH`)
  - `HeaderRowHeight` (not `HeaderRowH`)
  - `SmallColumnWidth` (not `SmallColW`)
  - `DefaultColumnWidth` (not `DefaultColW`)

#### Updated CATV/TVD Formatting Section
- **Location:** Lines 1133-1155
- **Changes:**
  - Added `$layout = Get-GridLayout $gridType` at STEP 5
  - Replaced `$FontSize` with `$layout.FontSize`
  - Replaced `$SmallColW` with `$layout.SmallColW`
  - Replaced `$DefaultColW` with `$layout.DefaultColW`
  - Replaced `$DefaultRowH` with `$layout.DefaultRowH`
  - Replaced `$HeaderRowH` with `$layout.HeaderRowH`

#### Updated Pasiones/Todo Formatting Section
- **Location:** Lines 1268-1325
- **Changes:**
  - Added `$layout = Get-GridLayout $gridType` after STEP 8
  - Replaced all global constants with `$layout` properties in:
    - Title row font size (lines 1287, 1302)
    - Title row height (lines 1289, 1302)
    - UsedRange font size (line 1309)
    - Small column width (line 1312)
    - Default column width (line 1319)
    - Default row height (line 1324)
    - Header row height (lines 1325-1326)

### 2. Settings/settings.json

- **Updated Structure:** Added `layout` section
- **Format:** JSON with string values (Settings UI expects strings)
- **Properties per grid type:**
  ```json
  "layout": {
    "CATV": {
      "fontSize": "14",
      "defaultRowHeight": "25",
      "headerRowHeight": "35",
      "smallColumnWidth": "7",
      "defaultColumnWidth": "37"
    },
    ... (same for other grid types)
  }
  ```

### 3. Settings UI.ps1 (Already Implemented)

- **Tab:** "Layout" tab in the Settings UI
- **Features:**
  - Grid type dropdown selector
  - 5 numeric input fields with validation:
    - FontSize: 8-72
    - DefaultRowHeight: 10-100
    - HeaderRowHeight: 10-100
    - SmallColumnWidth: 1-50
    - DefaultColumnWidth: 1-100
  - Save button updates FormatGrids.ps1 and settings.json

### 4. Test Script

- **Location:** `AI/Tests/Test-GridLayout.ps1`
- **Purpose:** Validates `Get-GridLayout()` function
- **Tests:**
  - Default layout retrieval for all grid types
  - Custom layout values for specific grids
  - Fallback to default layout for unknown grid types

## How It Works

### Architecture

```
User opens Settings UI
        ↓
User modifies layout settings for a grid type
        ↓
Settings UI saves to:
  - settings.json (JSON format)
  - FormatGrids.ps1 (PowerShell managed block)
        ↓
User runs FormatGrids.ps1 on Excel files
        ↓
Formatter detects grid type (Get-ExistingGridType)
        ↓
Formatter calls Get-GridLayout($gridType)
        ↓
Get-GridLayout() reads from $ManagedSettings.Layouts
        ↓
Formatter uses returned layout values for formatting
```

### Property Name Mapping

The system uses different property names in different places:

| Context | Property Name | Example |
|---------|--------------|---------|
| Settings UI (JSON) | camelCase | `defaultRowHeight` |
| Settings UI (PowerShell defaults) | camelCase | `defaultRowHeight` |
| Managed Settings Block | PascalCase | `DefaultRowHeight` |
| Get-GridLayout() return | Short form | `DefaultRowH` |
| Formatter code | Short form | `$layout.DefaultRowH` |

**Why?**
- Settings UI uses camelCase for JSON compatibility and user-friendly defaults
- Managed Block uses PascalCase (PowerShell convention)
- Formatter uses short form (established naming in existing code)

**Get-GridLayout() handles the conversion:**
```powershell
DefaultRowH  = [double]$layout.DefaultRowHeight  # Long → Short
```

## Testing Instructions

### 1. Test Get-GridLayout Function

```powershell
cd "C:\Users\mauhd\GridsV2-Debug\Grids V2"
powershell -ExecutionPolicy Bypass -File "AI\Tests\Test-GridLayout.ps1"
```

**Expected Output:** All grid types return default layout values; custom CATV values override defaults.

### 2. Test Settings UI

1. Open `Settings UI.ps1`
2. Navigate to the "Layout" tab
3. Select a grid type (e.g., "CATV")
4. Modify layout values:
   - FontSize: 16
   - DefaultRowHeight: 30
   - HeaderRowHeight: 40
   - SmallColumnWidth: 8
   - DefaultColumnWidth: 40
5. Click "Save Settings"
6. Verify `settings.json` was updated
7. Verify `FormatGrids.ps1` managed block was updated

### 3. Test Formatter with Real Files

**Setup:**
1. Prepare test Excel files for each grid type:
   - `GRILLA CATV MASTER Week 31.xlsx`
   - `GRILLA TVD MASTER Week 31.xlsx`
   - `GRILLA MASTER PASIONES LATAM.xlsx`
   - `GRILLA MASTER PASIONES US.xlsx`
   - `TODO NOVELAS.xlsx`
   - `REV_TV_GRID.xlsx`

2. Configure different layout settings for each grid type using Settings UI

**Run Test:**
```powershell
cd "C:\Users\mauhd\GridsV2-Debug\Grids V2"
.\FormatGrids.ps1
```

**Verify:**
- Open each formatted Excel file
- Check font size matches configured value
- Check row heights match configured values
- Check column widths match configured values
- Each grid type should have its own unique formatting

### 4. Test Fallback Behavior

1. Open `FormatGrids.ps1`
2. Remove a grid type from `$ManagedSettings.Layouts` (e.g., delete SPORTYNET)
3. Run formatter on a SportyNet file
4. Verify it uses default layout values (no crash)

## Validation Rules

Settings UI enforces these constraints:

| Property | Min | Max | Default |
|----------|-----|-----|---------|
| FontSize | 8 | 72 | 14 |
| DefaultRowHeight | 10 | 100 | 25 |
| HeaderRowHeight | 10 | 100 | 35 |
| SmallColumnWidth | 1 | 50 | 7 |
| DefaultColumnWidth | 1 | 100 | 37 |

## Version Information

- **FormatGrids.ps1:** Version updated to 2.0.3-dev
- **Last Modified:** 2026-07-31
- **Commit:** c0224cb (from user report)

## Known Limitations

1. **REV TV Grid:** Uses hardcoded column widths by header name (lines 977-994), not affected by layout settings. This is intentional - REV TV has a special structure.

2. **SportyNet Grid:** Has COM interaction issues (from user report), layout settings won't apply until this is fixed.

3. **Settings UI Property Names:** Must use camelCase in settings.json and PascalCase in managed block. The mapping is handled by Get-GridLayout().

## Troubleshooting

### Issue: Grid uses default values instead of custom values

**Check:**
1. Verify grid type name matches exactly (case-sensitive)
2. Check `settings.json` has correct structure
3. Check `FormatGrids.ps1` managed block was updated by Settings UI
4. Verify Get-GridLayout() is being called (add `Write-Host` debug output)

### Issue: Settings UI fails to save

**Check:**
1. Verify all values are within validation ranges
2. Check FormatGrids.ps1 is not open in another program
3. Verify Settings UI has write permissions to the folder
4. Check PowerShell syntax validation in Settings UI error message

### Issue: Formatter crashes with "Cannot index into a null array"

**Cause:** `$ManagedSettings.Layouts` doesn't exist or is null

**Fix:**
1. Run Settings UI and save settings (this creates the Layouts section)
2. Or manually add Layouts section to managed block in FormatGrids.ps1

## Next Steps

To complete the implementation:

1. ✅ Add Get-GridLayout() helper function
2. ✅ Update managed settings structure
3. ✅ Update CATV/TVD formatting code
4. ✅ Update Pasiones/Todo formatting code
5. ✅ Copy updated FormatGrids_V2.ps1 to FormatGrids.ps1
6. ⏳ Test with 6 working grid types
7. ✅ Settings UI already supports layout settings (user reported this is done)

**Remaining Work:**
- Test with actual Excel files for each grid type
- Verify different layout values apply correctly to each grid type
- Document user-facing instructions for configuring layout settings