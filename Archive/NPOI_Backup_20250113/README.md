# NPOI Excel Export Backup

**Backup Date:** January 13, 2025

## Purpose
This folder contains the original NPOI-based Excel export implementation before migrating to EPPlus.

## Original Configuration

### NuGet Package
- **Package:** NPOI
- **Version:** 2.5.6
- **License:** Apache License 2.0

### Dependencies
```xml
<PackageReference Include="NPOI" Version="2.5.6" />
```

### Key NPOI Classes Used
- `XSSFWorkbook` - Main workbook class
- `XSSFSheet` - Sheet manipulation
- `HSSFColor` - Color constants
- `XSSFSheetConditionalFormatting` - Conditional formatting
- `CellRangeAddress` - Cell range operations
- `BorderStyle`, `FillPattern`, `HorizontalAlignment` - Cell styling

### Files Backed Up
1. **ExportManager_NPOI_2.5.6.cs** - Complete Excel export implementation using NPOI

## Rollback Instructions

If you need to rollback to NPOI:

1. Copy `ExportManager_NPOI_2.5.6.cs` back to `ExportManager.cs`
2. Remove EPPlus package reference from `.csproj`
3. Add NPOI package reference:
   ```xml
   <PackageReference Include="NPOI" Version="2.5.6" />
   ```
4. Rebuild the solution

## Export Features Implemented
- FQQ IDNAC Designer format Excel export
- Multiple sheets (Parameters, IDNAC Table, Summary, Device Details, Calculations)
- Conditional formatting with color gradients
- Formula-based calculations
- Named ranges for dynamic updates
- Cell styling and formatting

## Notes
- NPOI 2.5.6 works well with .NET Framework 4.8
- No licensing fees required (Apache 2.0 license)
- Limited conditional formatting compared to EPPlus
- No true data bars support (used color gradient fallback)