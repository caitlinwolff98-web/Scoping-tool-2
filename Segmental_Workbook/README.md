# Bidvest Segmental Scoping Tool

## Overview

The **Bidvest Segmental Scoping Tool** is a comprehensive VBA-based Excel macro workbook designed for analyzing segmental financial reports. This tool processes segment-level data, creates aggregated tables, calculates percentages based on base amounts, and generates Power BI-ready outputs for audit scoping and analysis.

## Version Information

- **Version**: 1.0.0
- **Release Date**: November 2025
- **Compatibility**: Excel 2016 or later (Windows)

## Key Features

### 1. **Interactive Tab Categorization**
- Prompts users to categorize each worksheet as:
  - **Segment Summary Tab** - Contains summary data and base amounts
  - **Segment Tab** - Individual segment data tabs
  - **Uncategorized** - Ignored in analysis
- Automatically suggests tab names as segment names
- Validates that required tabs are present

### 2. **Intelligent Row 6 Analysis**
- Analyzes Row 6 in each segment tab to identify column markers
- Automatically excludes columns containing:
  - "Total"
  - "All journals"
  - "Base amounts"
- Only processes columns with pack data (Row 6 empty)

### 3. **Pack Name and Code Separation**
- Automatically separates pack names from pack codes
- Supports multiple formats:
  - "Pack Name (Code)"
  - "Pack Name - Code"
  - "Pack Name"

### 4. **Base Amounts Calculation**
- Prompts user to identify the "Base amounts" label in Summary tab Row 7
- Finds all columns containing the specified label
- Calculates total base amounts per FSLI by aggregating across base amounts columns
- Creates transparent calculation table showing the breakdown

### 5. **FullTable Generation**
- Creates a comprehensive table with:
  - Row 1: All FSLIs as column headers
  - Column A: Pack names
  - Column B: Pack codes
  - Column C: Segment names
  - Columns 4+: FSLI values
- Aggregates data from all segment tabs

### 6. **FullTablePercentage Generation**
- Converts all FSLI values to percentages
- Base 100% = Total base amounts from Summary tab
- Shows each pack's contribution to the total base amount per FSLI

### 7. **Power BI Integration**
Creates all necessary tables for Power BI dashboards:
- **ScopingControl** - Main scoping table with all pack/FSLI combinations
- **PackNumberCompanyTable** - Pack lookup table
- **FSLIKeyTable** - FSLI classification and metadata
- **SegmentList** - Segment information
- **ThresholdConfiguration** - Scoping threshold settings
- **ManualScopingAdjustments** - Manual override tracking
- **PowerBI_Metadata** - Run information and metadata

### 8. **Interactive Scoping Dashboard**
- Dynamic dashboard with real-time statistics
- Summary statistics display
- Threshold configuration controls
- Scoping by segment breakdown
- Scoping by FSLI breakdown
- One-click threshold application
- Manual scoping adjustments

### 9. **Threshold-Based Scoping**
- Configurable materiality, performance, and trivial thresholds
- Automatic scoping based on percentage thresholds
- Tracks scoping method (threshold vs manual)
- Supports manual overrides

## Architecture

### VBA Modules

The tool consists of the following VBA modules:

1. **ModConfig.bas**
   - Global constants and configuration
   - Utility functions
   - Validation functions

2. **ModTabCategorization.bas**
   - Tab categorization prompts
   - Segment name collection
   - Base amounts label identification

3. **ModDataProcessing.bas**
   - Row 6 analysis logic
   - FSLI extraction
   - Pack name/code separation
   - Base amounts calculation
   - Data validation

4. **ModTableGeneration.bas**
   - FullTable generation
   - FullTablePercentage generation
   - BaseAmountsCalculation table creation

5. **ModPowerBIIntegration.bas**
   - ScopingControl table generation
   - Pack Number Company table
   - FSLI Key table
   - Segment List table
   - Threshold Configuration table
   - Manual Scoping table
   - PowerBI Metadata sheet

6. **ModInteractiveDashboard.bas**
   - Dashboard creation and layout
   - Real-time data refresh
   - Threshold scoping logic
   - Scoping statistics

7. **ModMain.bas**
   - Main orchestration logic
   - Workflow coordination
   - Quick access procedures
   - Export and refresh functions

8. **ThisWorkbook.cls**
   - Workbook event handlers
   - Initialization logic
   - Quick Access Toolbar procedures

## Workflow

### Complete Analysis Workflow

```
1. User opens segmental workbook
   ↓
2. Runs RunSegmentalAnalysis macro
   ↓
3. Tool prompts for source workbook name
   ↓
4. Tool prompts to categorize each tab
   - Segment Summary Tab
   - Segment Tab (with segment name)
   - Uncategorized
   ↓
5. User identifies Summary tab
   ↓
6. Tool prompts for "Base amounts" label
   ↓
7. Tool validates base amounts columns exist
   ↓
8. Tool generates FullTable from all segments
   - Analyzes Row 6 to exclude totals/journals
   - Extracts pack names and codes from Row 8
   - Collects all unique FSLIs
   - Builds aggregated table
   ↓
9. Tool generates FullTablePercentage
   - Calculates base amounts per FSLI from Summary tab
   - Converts all values to percentages
   - Creates calculation transparency table
   ↓
10. Tool generates all Power BI tables
   - ScopingControl
   - PackNumberCompanyTable
   - FSLIKeyTable
   - SegmentList
   - ThresholdConfiguration
   - ManualScopingAdjustments
   - PowerBI_Metadata
   ↓
11. Tool creates results workbook
   ↓
12. User can launch Interactive Dashboard
   ↓
13. User applies threshold-based scoping
   ↓
14. User makes manual adjustments if needed
   ↓
15. User exports to Power BI
```

## Table Structures

### FullTable

| Pack Name | Pack Code | Segment | FSLI 1 | FSLI 2 | ... | FSLI N |
|-----------|-----------|---------|--------|--------|-----|--------|
| Pack A    | PA001     | Seg 1   | 1000   | 2000   | ... | 500    |
| Pack A    | PA001     | Seg 2   | 1500   | 1800   | ... | 600    |
| Pack B    | PB002     | Seg 1   | 800    | 1200   | ... | 300    |

### FullTablePercentage

Same structure as FullTable, but all FSLI values are percentages of the base amounts.

### ScopingControl

| Pack Code | Pack Name | Segment | FSLI | Amount | Percentage | Is Scoped | Scoping Method | Threshold Level | Manual Override | Comments |
|-----------|-----------|---------|------|--------|------------|-----------|----------------|-----------------|-----------------|----------|
| PA001     | Pack A    | Seg 1   | Rev  | 1000   | 5.2%       | Yes       | Threshold      | 2%              | No              |          |

### BaseAmountsCalculation

| FSLI | Column 1 | Column 2 | ... | Column N | Total Base Amount |
|------|----------|----------|-----|----------|-------------------|
| Rev  | 5000     | 8000     | ... | 2000     | 19250             |
| Cost | 3000     | 4500     | ... | 1500     | 12000             |

## Requirements

### System Requirements
- Microsoft Excel 2016 or later (Windows)
- VBA enabled
- Microsoft Scripting Runtime reference enabled

### Input Data Requirements

Your segment report workbook must have the following structure:

1. **Segment Summary Tab** (at least one)
   - Row 7: Contains "Base amounts" label (or custom label) identifying base amounts columns
   - Column B: FSLIs
   - FSLIs should be consistent across all tabs

2. **Segment Tabs** (at least one)
   - Row 6: Contains markers ("Total", "All journals", "Base amounts") or empty
     - Columns with markers are excluded from analysis
     - Columns with empty Row 6 are included
   - Row 8: Contains pack names and codes (format: "Pack Name (Code)")
   - Column B: FSLIs (matching Summary tab FSLIs)

## Installation Instructions

See [INSTALLATION.md](INSTALLATION.md) for detailed setup instructions.

## Usage Guide

See [USAGE_GUIDE.md](USAGE_GUIDE.md) for step-by-step usage instructions.

## Troubleshooting

### Common Issues

#### "Microsoft Scripting Runtime is not available"
**Solution**: Enable the reference in VBA
1. Open VBA Editor (Alt + F11)
2. Go to Tools > References
3. Check "Microsoft Scripting Runtime"
4. Click OK

#### "Could not identify Base amounts columns"
**Solution**: Verify the label
1. Check Row 7 in your Summary tab
2. Find the text that identifies base amounts columns
3. Enter the exact text when prompted (case-insensitive matching)

#### "No Segment tabs found"
**Solution**: Ensure proper categorization
1. At least one tab must be categorized as "Segment Tab"
2. Re-run the analysis and categorize tabs correctly

#### "Row 6 analysis excludes all columns"
**Solution**: Check Row 6 content
1. Verify Row 6 doesn't contain "Total", "All journals", or "Base amounts" in pack columns
2. Row 6 should be empty for pack data columns

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the usage guide
3. Check the VBA code comments for detailed logic
4. Contact the development team

## License

This tool is proprietary to Bidvest and is for internal use only.

## Changelog

### Version 1.0.0 (November 2025)
- Initial release
- Full segment analysis capability
- Interactive dashboard
- Power BI integration
- Threshold-based scoping
- Manual scoping adjustments

## Credits

Developed for Bidvest by Claude (Anthropic)
Based on requirements for segment-level audit scoping and analysis.

---

**Last Updated**: November 21, 2025
