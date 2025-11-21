# Usage Guide

## Bidvest Segmental Scoping Tool - Complete User Guide

This guide provides step-by-step instructions for using the Bidvest Segmental Scoping Tool to analyze segmental financial reports.

## Table of Contents

1. [Quick Start](#quick-start)
2. [Preparing Your Segment Report](#preparing-your-segment-report)
3. [Running the Full Analysis](#running-the-full-analysis)
4. [Understanding the Output](#understanding-the-output)
5. [Using the Interactive Dashboard](#using-the-interactive-dashboard)
6. [Applying Threshold Scoping](#applying-threshold-scoping)
7. [Making Manual Adjustments](#making-manual-adjustments)
8. [Exporting to Power BI](#exporting-to-power-bi)
9. [Advanced Features](#advanced-features)
10. [Troubleshooting](#troubleshooting)

---

## Quick Start

### 5-Minute Quick Start

1. Open your segment report workbook in Excel
2. Open the Bidvest Segmental Scoping Tool workbook
3. Press **Alt + F8**, select `RunSegmentalAnalysis`, click **Run**
4. Follow the prompts to categorize tabs and identify base amounts
5. Review the generated output tables

---

## Preparing Your Segment Report

Before running the analysis, ensure your segment report workbook has the correct structure.

### Required Structure

#### Segment Summary Tab
- **Row 7**: Must contain a label identifying "Base amounts" columns (e.g., "Base amounts", "Base", "Original", etc.)
- **Column B**: Must contain FSLIs
- Example:

| Row | A | B (FSLI) | C | D (Base amounts) | E (Base amounts) | F (All journals) | G (Total) |
|-----|---|----------|---|------------------|------------------|------------------|-----------|
| 7   |   |          |   | Base amounts     | Base amounts     | All journals     | Total     |
| 8   |   | Revenue  |   | 5000             | 8000             | 2000             | 15000     |

#### Segment Tabs
- **Row 6**: Contains markers or is empty
  - Empty = include this column
  - "Total" / "All journals" / "Base amounts" = exclude this column
- **Row 8**: Contains pack names and codes (format: "Pack Name (Code)" or "Pack Name - Code")
- **Column B**: Contains FSLIs (should match Summary tab FSLIs)

Example:

| Row | A | B (FSLI)  | C                  | D                | E (Total) | F (All journals) |
|-----|---|-----------|--------------------|------------------|-----------|------------------|
| 6   |   |           |                    |                  | Total     | All journals     |
| 8   |   |           | Pack A (PA001)     | Pack B (PB002)   |           |                  |
| 9   |   | Revenue   | 1000               | 800              | 1800      | 200              |
| 10  |   | Expenses  | 600                | 500              | 1100      | 150              |

### Checklist Before Running

- [ ] All segment tabs have FSLIs in Column B
- [ ] Row 6 contains appropriate markers (or is empty for pack columns)
- [ ] Row 8 contains pack information with codes
- [ ] Summary tab Row 7 has base amounts label
- [ ] All tabs have consistent FSLI naming
- [ ] Workbook is saved and backup created

---

## Running the Full Analysis

### Step-by-Step Process

#### Step 1: Launch the Analysis

1. Open the **Bidvest Segmental Scoping Tool.xlsm** workbook
2. Ensure your segment report workbook is also open
3. Press **Alt + F8** to open the Macro dialog
4. Select **RunSegmentalAnalysis**
5. Click **Run**

#### Step 2: Select Source Workbook

A dialog will appear listing all open workbooks.

1. Enter the name of your segment report workbook
   - Example: `Bidvest Segments 2025.xlsx`
   - You can enter with or without the extension
2. Click **OK**

> **Tip**: The workbook name is case-insensitive

#### Step 3: Categorize Tabs

For each worksheet in your source workbook, you'll be prompted:

**"How should the tab 'TabName' be categorized?"**

Choose:
- **YES** - This is the **Segment Summary Tab**
  - Only select ONE tab as Summary
  - This is the tab with base amounts in Row 7

- **NO** - This is a **Segment Tab**
  - You'll then be prompted to enter the segment name
  - Default suggestion: The tab name itself
  - Edit if needed (e.g., change "Seg1" to "Automotive Division")

- **CANCEL** - **Uncategorized** (skip this tab)
  - Use for working tabs, notes, etc.

**Example Categorization:**
```
Tab "Summary" → YES (Segment Summary)
Tab "Automotive" → NO → Enter segment name: "Automotive Division"
Tab "Services" → NO → Enter segment name: "Services Division"
Tab "Notes" → CANCEL (Uncategorized)
```

> **Important**: You must categorize at least:
> - 1 Segment Summary Tab
> - 1 Segment Tab

#### Step 4: Identify Summary Tab

You'll be prompted to enter the name of the Segment Summary tab.

1. Enter the exact name of your summary tab
   - Example: `Summary`
2. Click **OK**

#### Step 5: Identify Base Amounts Label

The tool will scan Row 7 of the Summary tab and show you potential labels.

**Prompt**: "Please enter the exact text that identifies the 'Base amounts' columns"

1. The tool will display labels found in Row 7
2. Enter the text that identifies base amounts columns
   - Example: `Base amounts`
   - Matching is case-insensitive
   - The tool will find all columns containing this text
3. Click **OK**

**Example:**
```
Found in Row 7: "Base amounts", "Base amounts", "All journals", "Total"
Enter label: Base amounts
Result: Columns with "Base amounts" will be used for base calculation
```

#### Step 6: Wait for Processing

The tool will now:
1. ✓ Generate FullTable from all segment tabs
2. ✓ Generate FullTablePercentage using base amounts
3. ✓ Create BaseAmountsCalculation table
4. ✓ Generate all Power BI tables
5. ✓ Apply initial scoping rules

**Progress indicators:**
- Watch the status bar at the bottom of Excel
- Messages appear in the Immediate Window (Ctrl + G in VBA Editor)

#### Step 7: Review Completion Message

When complete, you'll see a success message:

```
Segmental analysis completed successfully!

Results saved to: Book1.xlsx

Total Segments: 3
Total FSLIs: 15
Total Packs: 12
```

Click **OK** to close the message.

---

## Understanding the Output

The tool creates a new workbook with multiple worksheets:

### Output Worksheets

#### 1. Instructions
- Overview of the output workbook
- List of all generated tables
- Run information (date, version, etc.)

#### 2. FullTable
The main aggregated data table.

**Structure:**
- **Column A**: Pack Name
- **Column B**: Pack Code
- **Column C**: Segment
- **Columns 4+**: FSLIs (one column per FSLI)

**Example:**
| Pack Name | Pack Code | Segment    | Revenue | Expenses | Assets |
|-----------|-----------|------------|---------|----------|--------|
| Pack A    | PA001     | Automotive | 1000    | 600      | 5000   |
| Pack A    | PA001     | Services   | 1500    | 800      | 6000   |
| Pack B    | PB002     | Automotive | 800     | 500      | 4000   |

**Use this table to:**
- See all segment data in one place
- Identify which packs appear in which segments
- Compare values across FSLIs

#### 3. FullTablePercentage
Same structure as FullTable, but values are percentages.

**Calculation:**
```
Percentage = (Pack FSLI Value / Total Base Amount for FSLI) × 100
```

**Example:**
| Pack Name | Pack Code | Segment    | Revenue | Expenses |
|-----------|-----------|------------|---------|----------|
| Pack A    | PA001     | Automotive | 5.2%    | 5.0%     |
| Pack A    | PA001     | Services   | 7.8%    | 6.7%     |

**Use this table to:**
- Identify significant packs (high percentages)
- Apply materiality thresholds
- Prioritize audit focus areas

#### 4. BaseAmountsCalculation
Shows how base amounts were calculated for transparency.

**Structure:**
- **Column A**: FSLI
- **Columns B-N**: Each base amounts column from Summary tab
- **Last Column**: Total Base Amount (sum of all base amounts columns)

**Example:**
| FSLI     | Column D | Column E | Total Base Amount |
|----------|----------|----------|-------------------|
| Revenue  | 5000     | 8000     | 13000             |
| Expenses | 3000     | 4500     | 7500              |

**Use this table to:**
- Verify base amounts calculation
- Understand which columns were included
- Audit trail and documentation

#### 5. ScopingControl
The main table for Power BI and scoping decisions.

**Columns:**
- Pack Code, Pack Name, Segment, FSLI
- Amount, Percentage
- Is Scoped (Yes/No)
- Scoping Method (Threshold/Manual/None)
- Threshold Level
- Manual Override (Yes/No)
- Comments

**Use this table to:**
- Track scoping decisions
- Apply threshold rules
- Make manual adjustments
- Feed Power BI dashboard

#### 6. PackNumberCompanyTable
Lookup table for packs.

**Columns:**
- Pack Code, Pack Name, Company Name, Company Code, Segment, Active

#### 7. FSLIKeyTable
Lookup table for FSLIs.

**Columns:**
- FSLI, FSLI Description, Category, Financial Statement, Sort Order

**Categories automatically assigned:**
- Revenue, Expense, Asset, Liability, Equity, Other

#### 8. SegmentList
List of all segments.

**Columns:**
- Segment Name, Segment Code, Tab Name, Active

#### 9. ThresholdConfiguration
Scoping threshold settings.

**Default thresholds:**
- Overall Materiality: 5%
- Performance Materiality: 2%
- Trivial Threshold: 0.5%

**Edit these values** to change scoping behavior.

#### 10. ManualScopingAdjustments
Track manual overrides to scoping.

**Columns:**
- Pack Code, Segment, FSLI, Scope In/Out, Reason, Adjusted By, Adjustment Date

#### 11. PowerBI_Metadata
Run information and metadata.

**Contains:**
- Tool version
- Generated date
- Base amounts label used
- Total segments, FSLIs, packs

---

## Using the Interactive Dashboard

### Launching the Dashboard

1. In the output workbook, press **Alt + F8**
2. Select **ShowInteractiveDashboard**
3. Click **Run**

OR

1. Go to the Developer tab
2. Click **Macros**
3. Select **ShowInteractiveDashboard**
4. Click **Run**

### Dashboard Layout

#### Summary Statistics Section
- **Total Items**: Total pack/FSLI combinations
- **Items Scoped In**: Currently scoped for audit
- **Items Scoped Out**: Currently not scoped
- **Scoping Coverage %**: Percentage scoped in
- **Total Segments/FSLIs/Packs**: Data summary

#### Threshold Configuration Section
- **Materiality %**: Overall materiality threshold
- **Performance %**: Performance materiality (workable threshold)
- **Trivial %**: Trivial threshold

**To change thresholds:**
1. Click on the percentage cell
2. Enter new value (e.g., 0.03 for 3%)
3. Press Enter
4. Apply thresholds using the action button

#### Scoping by Segment Section
Table showing scoping statistics per segment:
- Segment name
- Total items in segment
- Items scoped in
- Items scoped out
- Coverage percentage

#### Scoping by FSLI Section
Table showing scoping statistics per FSLI:
- FSLI name
- Total items
- Items scoped in
- Coverage percentage

### Dashboard Actions

#### [Apply Thresholds]
- Applies threshold-based scoping to all items
- Uses Performance % threshold
- Items >= threshold are scoped in
- Items < threshold are scoped out

#### [Clear Scoping]
- Clears all scoping decisions
- Resets all items to "Not Scoped"
- Useful for starting fresh

#### [Export Results]
- Prepares data for Power BI export
- Validates all required tables exist

#### [Refresh Data]
- Refreshes dashboard statistics
- Updates all calculations
- Use after making manual changes

---

## Applying Threshold Scoping

### Automatic Threshold Scoping

#### Method 1: Via Dashboard

1. Launch the dashboard (`ShowInteractiveDashboard`)
2. Verify/adjust threshold percentages in the Threshold Configuration section
3. Click **[Apply Thresholds]** button or run macro `ApplyThresholdScoping`
4. Review the confirmation message

#### Method 2: Via Macro

1. Press **Alt + F8**
2. Select **ApplyThresholdScoping**
3. Click **Run**

### How Threshold Scoping Works

1. The tool reads the **Performance %** threshold (default: 2%)
2. For each row in ScopingControl:
   - If Percentage >= Performance %, mark as "Scoped In"
   - If Percentage < Performance %, mark as "Scoped Out"
3. Updates columns:
   - **Is Scoped**: "Yes" or "No"
   - **Scoping Method**: "Threshold"
   - **Threshold Level**: The percentage used

### Customizing Thresholds

Edit values in the **ThresholdConfiguration** table:

1. Go to the **ThresholdConfiguration** worksheet
2. Edit the percentage values:
   - Overall Materiality: Higher level (e.g., 5%)
   - Performance Materiality: Working threshold (e.g., 2%)
   - Trivial Threshold: Very small items (e.g., 0.5%)
3. Run **ApplyThresholdScoping** to apply new thresholds

**Common threshold scenarios:**

| Scenario | Materiality | Performance | Trivial |
|----------|-------------|-------------|---------|
| Conservative | 3% | 1.5% | 0.3% |
| Standard | 5% | 2% | 0.5% |
| Aggressive | 8% | 3% | 1% |

---

## Making Manual Adjustments

Sometimes you need to override threshold scoping decisions.

### Manual Scope In/Out

1. Open the **ScopingControl** worksheet
2. Find the row for the pack/segment/FSLI combination
3. In the **Is Scoped** column (Column G), change:
   - To "Yes" to scope in
   - To "No" to scope out
4. In the **Manual Override** column (Column J), enter "Yes"
5. In the **Comments** column (Column K), add a reason (optional but recommended)

**Example:**
| Pack Code | FSLI | Is Scoped | Manual Override | Comments |
|-----------|------|-----------|-----------------|----------|
| PA001 | Rev | Yes | Yes | High risk area - scope in regardless of threshold |

### Tracking Manual Adjustments

The **ManualScopingAdjustments** table automatically tracks:
- Which items were manually adjusted
- Who made the adjustment
- When it was made
- The reason

**To add to tracking table:**

1. Go to **ManualScopingAdjustments** worksheet
2. Add a new row with:
   - Pack Code
   - Segment
   - FSLI
   - Scope In/Out
   - Reason
   - Adjusted By (your name)
   - Adjustment Date

### Best Practices for Manual Adjustments

✓ **DO:**
- Document the reason for manual adjustments
- Review manual adjustments with your team
- Keep a record in ManualScopingAdjustments table

✗ **DON'T:**
- Make manual adjustments without documentation
- Override thresholds without valid business reason
- Forget to update Comments column

---

## Exporting to Power BI

### Preparation

Before exporting, ensure:
1. All scoping decisions are finalized
2. Manual adjustments are documented
3. All required tables exist

### Export Process

#### Method 1: Via Macro

1. Press **Alt + F8**
2. Select **ExportToPowerBI**
3. Click **Run**
4. Review the confirmation message

The tool will verify all required tables exist:
- ScopingControl ✓
- PackNumberCompanyTable ✓
- FSLIKeyTable ✓
- SegmentList ✓
- ThresholdConfiguration ✓
- ManualScopingAdjustments ✓

#### Method 2: Save for Power BI

1. Save the output workbook to a location accessible to Power BI
2. Note the file path (you'll need it in Power BI)

### Connecting Power BI

1. Open Power BI Desktop
2. Click **Get Data** > **Excel**
3. Navigate to the output workbook
4. Select the following tables:
   - [x] ScopingControl
   - [x] PackNumberCompanyTable
   - [x] FSLIKeyTable
   - [x] SegmentList
   - [x] ThresholdConfiguration
   - [x] FullTable (optional - for detailed analysis)
   - [x] FullTablePercentage (optional)
5. Click **Load**

### Creating Relationships in Power BI

After loading, create relationships:

1. Go to **Model** view
2. Create relationships:
   - ScopingControl[Pack Code] → PackNumberCompanyTable[Pack Code]
   - ScopingControl[FSLI] → FSLIKeyTable[FSLI]
   - ScopingControl[Segment] → SegmentList[Segment Name]

### Sample Power BI Visualizations

**Recommended visualizations:**
1. **Scoping Summary Card**
   - Total items scoped in
   - Total items scoped out
   - Coverage percentage

2. **Scoping by Segment** (Bar Chart)
   - X-axis: Segment
   - Y-axis: Count of items scoped in
   - Legend: Is Scoped

3. **Scoping by FSLI** (Stacked Bar)
   - X-axis: FSLI
   - Y-axis: Count
   - Legend: Is Scoped

4. **Top Packs by Percentage** (Table)
   - Pack Name
   - Segment
   - FSLI
   - Percentage
   - Sorted by Percentage descending

5. **Threshold Analysis** (Scatter Plot)
   - X-axis: FSLI
   - Y-axis: Percentage
   - Size: Amount
   - Color: Is Scoped
   - Reference line at Performance threshold

---

## Advanced Features

### Regenerate Specific Tables

#### Regenerate FullTable Only

1. Press **Alt + F8**
2. Select **RegenerateFullTable**
3. Click **Run**
4. Select source workbook when prompted

#### Regenerate Power BI Tables Only

1. Press **Alt + F8**
2. Select **RegeneratePowerBITables**
3. Click **Run**

### Refresh All Data

To refresh all data from source:

1. Press **Alt + F8**
2. Select **RefreshAllData**
3. Click **Run**
4. Confirm when prompted

This will re-run the entire analysis.

### Debugging and Logging

To view detailed logs:

1. Press **Alt + F11** to open VBA Editor
2. Press **Ctrl + G** to open Immediate Window
3. Run any macro
4. View detailed log messages in Immediate Window

**Log messages include:**
- Timestamps
- Processing steps
- Row counts
- Warnings and errors

---

## Troubleshooting

### Common Issues and Solutions

#### Issue: "Could not identify Base amounts columns in Summary tab"

**Cause**: The label entered doesn't match any text in Row 7

**Solution**:
1. Open the Summary tab
2. Look at Row 7
3. Find the exact text in the base amounts columns
4. Re-run and enter that exact text (case doesn't matter)

#### Issue: "No Segment tabs found"

**Cause**: No tabs were categorized as Segment Tabs

**Solution**:
1. Re-run the analysis
2. When prompted, select **NO** for at least one tab
3. Enter a segment name when prompted

#### Issue: "FSLIs don't match between Summary and Segment tabs"

**Cause**: FSLI names are different across tabs

**Solution**:
1. Standardize FSLI names across all tabs
2. Ensure exact spelling (case doesn't matter)
3. Remove extra spaces

#### Issue: "Row 6 excludes all columns"

**Cause**: Row 6 contains markers in all columns

**Solution**:
1. Check Row 6 in your segment tabs
2. Pack data columns should have Row 6 empty
3. Only total/journal columns should have markers

#### Issue: "Percentages are incorrect"

**Cause**: Wrong base amounts columns were used

**Solution**:
1. Check the **BaseAmountsCalculation** table
2. Verify the correct columns were included
3. Re-run and enter the correct base amounts label

#### Issue: "Some packs are missing from FullTable"

**Cause**: Pack information might be in excluded columns

**Solution**:
1. Check Row 6 - ensure pack columns don't have markers
2. Check Row 8 - ensure pack names and codes exist
3. Verify pack name format: "Pack Name (Code)"

### Getting Help

If you encounter other issues:

1. **Check the error message** - It often contains helpful information
2. **Review the Immediate Window** - Shows detailed logging (Ctrl + G in VBA Editor)
3. **Verify prerequisites** - Check INSTALLATION.md
4. **Test with sample data** - Isolate whether it's a data issue
5. **Check VBA compilation** - In VBA Editor: Debug > Compile VBAProject

---

## Best Practices

### Data Preparation
- ✓ Always create a backup of your source workbook
- ✓ Ensure consistent FSLI naming across all tabs
- ✓ Use consistent pack code formats
- ✓ Remove any circular formulas in source data

### Running Analysis
- ✓ Close unnecessary workbooks to avoid confusion
- ✓ Save your work before running long processes
- ✓ Review categorization carefully - it affects all outputs
- ✓ Double-check the base amounts label

### Scoping Decisions
- ✓ Start with threshold-based scoping
- ✓ Document all manual adjustments
- ✓ Review scoping coverage percentages
- ✓ Validate with your team before finalizing

### Power BI Integration
- ✓ Save output workbook in a stable location
- ✓ Use consistent file names for easy refresh
- ✓ Test relationships in Power BI
- ✓ Create bookmarks for different scoping scenarios

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Alt + F8 | Open Macros dialog |
| Alt + F11 | Open VBA Editor |
| Ctrl + G | Open Immediate Window (in VBA Editor) |
| F5 | Run current procedure (in VBA Editor) |
| Ctrl + S | Save workbook |

---

## Appendix: Sample Workflow

### Complete End-to-End Workflow

1. **Prepare** (10 minutes)
   - Back up source workbook
   - Verify structure (Row 6, 7, 8, Column B)
   - Standardize FSLI names

2. **Run Analysis** (5 minutes)
   - Open source and tool workbooks
   - Run `RunSegmentalAnalysis`
   - Categorize tabs (Summary + Segments)
   - Identify base amounts label
   - Wait for processing

3. **Review Output** (15 minutes)
   - Check FullTable structure
   - Verify FullTablePercentage calculations
   - Review BaseAmountsCalculation
   - Validate ScopingControl data

4. **Apply Scoping** (10 minutes)
   - Launch Interactive Dashboard
   - Set threshold percentages
   - Apply threshold scoping
   - Review scoping coverage

5. **Manual Adjustments** (20 minutes)
   - Review high-risk areas
   - Scope in items as needed
   - Document reasons in Comments
   - Update ManualScopingAdjustments

6. **Finalize** (10 minutes)
   - Review final scoping statistics
   - Save output workbook
   - Export to Power BI
   - Create Power BI dashboard

**Total Time: ~70 minutes** (first time; faster subsequently)

---

**Usage Guide Version**: 1.0.0
**Last Updated**: November 21, 2025
