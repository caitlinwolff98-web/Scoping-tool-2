# Installation Guide

## Bidvest Segmental Scoping Tool - Installation Instructions

This guide will walk you through setting up the Bidvest Segmental Scoping Tool in Excel.

## Prerequisites

Before you begin, ensure you have:

- [ ] Microsoft Excel 2016 or later (Windows)
- [ ] Administrator rights to enable macros and VBA
- [ ] Access to the segmental report workbooks you want to analyze

## Installation Steps

### Step 1: Enable Developer Tab in Excel

The Developer tab is required to import VBA modules.

1. Open Excel
2. Click **File** > **Options**
3. Click **Customize Ribbon**
4. In the right pane, check the box for **Developer**
5. Click **OK**

### Step 2: Create a New Excel Workbook

1. Open Excel
2. Create a new blank workbook
3. Save it as:
   - **File name**: `Bidvest Segmental Scoping Tool.xlsm`
   - **Save as type**: Excel Macro-Enabled Workbook (*.xlsm)
4. Choose a location on your computer

### Step 3: Enable VBA References

1. Press **Alt + F11** to open the VBA Editor
2. Go to **Tools** > **References**
3. Check the following references (scroll to find them):
   - [x] **Visual Basic For Applications**
   - [x] **Microsoft Excel 16.0 Object Library**
   - [x] **OLE Automation**
   - [x] **Microsoft Office 16.0 Object Library**
   - [x] **Microsoft Scripting Runtime** ⭐ IMPORTANT
4. Click **OK**

> **Note**: If you don't see "Microsoft Scripting Runtime", try these steps:
> - Click **Browse**
> - Navigate to `C:\Windows\System32\scrrun.dll`
> - Select it and click **Open**

### Step 4: Import VBA Modules

You need to import all the VBA module files and class modules provided.

#### For each .bas file (Standard Modules):

1. In the VBA Editor (Alt + F11)
2. Right-click on **VBAProject** (your workbook name)
3. Select **Import File...**
4. Navigate to the folder containing the .bas files
5. Select the file and click **Open**

**Import these .bas files in order:**

1. `ModConfig.bas`
2. `ModTabCategorization.bas`
3. `ModDataProcessing.bas`
4. `ModTableGeneration.bas`
5. `ModPowerBIIntegration.bas`
6. `ModInteractiveDashboard.bas`
7. `ModMain.bas`

#### For each .cls file (Class Modules):

**IMPORTANT**: Use the _IMPORT.cls versions for importing!

1. In the VBA Editor (Alt + F11)
2. Right-click on **VBAProject** (your workbook name)
3. Select **Import File...**
4. Navigate to the folder containing the .cls files
5. Select the file and click **Open**

**Import these .cls files:**

1. `clsColumnAnalysis_IMPORT.cls` ⭐ (Use this version, not the regular .cls)
2. `clsPackInfo_IMPORT.cls` ⭐ (Use this version, not the regular .cls)

> **Why _IMPORT versions?** The _IMPORT.cls files have `VB_Creatable = True` which allows the classes to be instantiated with `New` and stored in Collections. The regular .cls files were generated with VB_Creatable = False which causes the compile error.

#### For the ThisWorkbook class module:

**IMPORTANT**: Do NOT import ThisWorkbook.cls as a file! Follow these steps instead:

1. In the VBA Editor, double-click **ThisWorkbook** in the Project Explorer (under Microsoft Excel Objects)
2. Delete all existing code in the ThisWorkbook code window
3. Open `ThisWorkbook_CODE_ONLY.txt` in a text editor
4. Copy **ALL** the code from that file
5. Paste it into the ThisWorkbook code window

> **Why not import?** The .cls file contains VBA metadata headers (VERSION, BEGIN, END, Attribute) that Excel generates automatically. Importing or copying these lines causes "Invalid outside procedure" errors. Use the ThisWorkbook_CODE_ONLY.txt file instead, which contains only the actual VBA code.

### Step 5: Verify Installation

1. In the VBA Editor, you should see all modules listed under **Modules**:
   - ModConfig
   - ModDataProcessing
   - ModInteractiveDashboard
   - ModMain
   - ModPowerBIIntegration
   - ModTabCategorization
   - ModTableGeneration

2. You should see class modules listed under **Class Modules**:
   - clsColumnAnalysis
   - clsPackInfo

3. You should see **ThisWorkbook** under **Microsoft Excel Objects**

### Step 6: Enable Macros

1. Close the VBA Editor
2. Close and reopen the workbook
3. If you see a **Security Warning** banner, click **Enable Content**

### Step 7: Configure Macro Security (One-Time Setup)

To avoid security warnings every time:

1. Click **File** > **Options**
2. Click **Trust Center**
3. Click **Trust Center Settings**
4. Click **Macro Settings**
5. Select **Enable all macros** (for development) OR
6. Select **Disable all macros with notification** (recommended for production)
7. Check **Trust access to the VBA project object model**
8. Click **OK** and **OK** again

### Step 8: Add to Trusted Locations (Recommended)

For easier access without security warnings:

1. Click **File** > **Options**
2. Click **Trust Center** > **Trust Center Settings**
3. Click **Trusted Locations**
4. Click **Add new location**
5. Browse to the folder where you saved the workbook
6. Check **Subfolders of this location are also trusted** (if needed)
7. Click **OK**

### Step 9: Test the Installation

1. Press **Alt + F8** to open the Macro dialog
2. You should see macros including:
   - `RunSegmentalAnalysis`
   - `ShowInteractiveDashboard`
   - `ApplyThresholdScoping`
   - etc.
3. Don't run them yet - just verify they appear

### Step 10: Create Quick Access Shortcuts (Optional)

For easier access, add macros to the Quick Access Toolbar:

1. Click the dropdown arrow on the Quick Access Toolbar (top left)
2. Select **More Commands**
3. In the **Choose commands from** dropdown, select **Macros**
4. Select a macro (e.g., `RunSegmentalAnalysis`)
5. Click **Add >>**
6. Repeat for other frequently used macros:
   - `ShowInteractiveDashboard`
   - `ApplyThresholdScoping`
   - `ExportToPowerBI`
7. Click **OK**

## Verification Checklist

Before proceeding to use the tool, verify:

- [x] All 7 modules are imported
- [x] ThisWorkbook class module is updated
- [x] Microsoft Scripting Runtime reference is enabled
- [x] Macros are enabled
- [x] No compile errors (in VBA Editor, click **Debug** > **Compile VBAProject**)

## Troubleshooting Installation Issues

### Issue: "Only user-defined types defined in public object modules..." error

**Cause**: You imported the wrong .cls files (without VB_Creatable = True)

**Solution**:
1. Open VBA Editor (Alt + F11)
2. In Project Explorer, expand **Class Modules**
3. Right-click on **clsColumnAnalysis** and select **Remove clsColumnAnalysis**
4. Click **No** when asked to export
5. Right-click on **clsPackInfo** and select **Remove clsPackInfo**
6. Click **No** when asked to export
7. Now import the correct files:
   - Right-click on **VBAProject** > **Import File...**
   - Select `clsColumnAnalysis_IMPORT.cls` and click Open
   - Right-click on **VBAProject** > **Import File...**
   - Select `clsPackInfo_IMPORT.cls` and click Open
8. Compile (Debug > Compile VBAProject)

> **The Fix**: The _IMPORT.cls files have `Attribute VB_Creatable = True` which allows `Set obj = New ClassName` to work properly!

### Issue: "Invalid outside procedure" in ThisWorkbook

**Cause**: You copied the VERSION/BEGIN/END/Attribute header lines from ThisWorkbook.cls

**Solution**:
1. Open VBA Editor (Alt + F11)
2. Double-click **ThisWorkbook** in the Project Explorer
3. Press Ctrl+A to select all code
4. Press Delete to clear everything
5. Open `ThisWorkbook_CODE_ONLY.txt` (not .cls!)
6. Copy all code from that file
7. Paste into ThisWorkbook code window
8. Save and compile (Debug > Compile VBAProject)

> **Key Point**: Never copy lines starting with "VERSION", "BEGIN", "END", "Attribute VB_" - these are metadata that Excel generates automatically!

### Issue: "Compile error: Can't find project or library"

**Solution**:
1. Open VBA Editor (Alt + F11)
2. Go to Tools > References
3. Look for any reference marked as **MISSING**
4. Uncheck the missing reference
5. Enable **Microsoft Scripting Runtime**
6. Click OK
7. Try compiling again (Debug > Compile VBAProject)

### Issue: "Import fails" or "File not found"

**Solution**:
1. Ensure all .bas files are in the same folder
2. Check file extensions (should be .bas, not .bas.txt)
3. Ensure you have read permissions on the files

### Issue: "Macro not appearing in list"

**Solution**:
1. Open VBA Editor
2. Find ModMain module
3. Verify `RunSegmentalAnalysis` procedure exists
4. Check that it's declared as `Public Sub`
5. Compile the project (Debug > Compile VBAProject)

### Issue: "Permission denied" errors

**Solution**:
1. Run Excel as Administrator (right-click Excel > Run as Administrator)
2. Retry the installation
3. Ensure antivirus is not blocking VBA

## Post-Installation

After successful installation:

1. **Save the workbook** (Ctrl + S)
2. **Create a backup copy** in a safe location
3. **Read the Usage Guide** (USAGE_GUIDE.md) to learn how to use the tool
4. **Test with sample data** before using with real data

## Updating the Tool

If you receive updated module files:

1. Open the VBA Editor (Alt + F11)
2. Right-click the module you want to update
3. Select **Remove [ModuleName]**
4. Click **No** when asked to export (unless you want a backup)
5. Import the new module file using **File** > **Import File**
6. Save the workbook

## Uninstallation

To remove the tool:

1. Open the VBA Editor (Alt + F11)
2. For each module:
   - Right-click the module
   - Select **Remove [ModuleName]**
   - Click **No** to export prompt
3. Delete the workbook file

---

## Support

If you encounter issues during installation:

1. Check the troubleshooting section above
2. Verify all prerequisites are met
3. Ensure Excel version compatibility (2016+)
4. Review the README.md for additional information

---

**Installation Guide Version**: 1.0.0
**Last Updated**: November 21, 2025
