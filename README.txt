DM COMPARE TOOL v3.1 |  Excel Ribbon Add-in
====================================

WHAT'S INCLUDED
- modDMCompare.bas           Complete VBA comparison engine / Ribbon macro
- frmFileCompare_code.txt   VBA UserForm event code
- Build_and_Install_DM_Compare.ps1   Windows Excel COM builder and installer
- Install_DM_Compare.cmd    Double-click launcher for the installer
- File_A_Sample.xlsx        Sample original data; header row 1, data row 2
- File_B_Sample.xlsx        Sample updated data; header row 3, data row 5


INSTALLATION (WINDOWS DESKTOP EXCEL REQUIRED)
1. Close all Excel windows
2. Open Excel > File > Options > Trust Center > Trust Center Settings >
   Macro Settings > check "Trust access to the VBA project object model".
   This is required only while the installer creates the add-in.
3. Close Excel, then double-click Install_DM_Compare.cmd.
4. Look for "SUCCESS" in the installer window. The actual .xlam file is
   generated locally in %APPDATA%\Microsoft\AddIns\DM Compare Tool.xlam.
5. Open Excel > File > Option > Customized Ribbon > Choose Commands "Popular commands" to "Macro" > Create Tab > Add in the tab "DM_CompareFiles".  
6. Open Excel. Use the "DM Compare" ribbon tab > "Compare Files" to compare.

Reminder:
Test FileB row header start at 3 and data start at 4.


SECURITY
- The setup script runs only local source code and Microsoft Excel automation.
- Review the included .ps1 and VBA source before installing if desired.
- Do not turn off Protected View globally or enable all macros globally.
- Depending on your organization, macros/add-ins may require IT approval.

USAGE
1. On Excel's "DM Compare" tab, click "Compare Files".
2. Browse File A and choose its worksheet.
3. Enter File A header row and first data row.
4. Browse File B and choose its worksheet.
5. Enter File B header row and first data row.
6. Click "Refresh headers" if needed, then select ONE common Unique ID.
7. Click "Compare Files". A NEW workbook opens with:
   - Total Data Rows - File A
   - Total Data Rows - File B
   - Total Mismatches (individual field differences)
   - Total Runtime (actual seconds, rounded to 2 decimals)
   - File A value / File B value / Correct or Mismatched status per field
   - Red font for mismatched values and status; green font for Correct
   - Rows existing in only one file flagged as Only in File A / Only in File B
8. Save the results workbook as .xlsx wherever you prefer.

TEST FILES
- File A: Sheet 'Employees', Header Row 1, Data Starting Row 2.
- File B: Sheet 'Updated Employees', Header Row 3, Data Starting Row 5.
- Select the common Unique ID 'Employee ID'.
- Expected: File A rows 4, File B rows 4, Mismatches 3.
  There is also one File A-only and one File B-only ID (not counted as a
  field mismatch). Both source files should remain unchanged.

NOTES
- Match headers ignore capitalization and leading/trailing spaces.
- Unique IDs are matched based on their cell value (whitespace trimmed).
- Text matches are case-sensitive. Blank and empty strings compare equal.
- Formula cells compare by resulting value, not formula text.
- Missing source records are flagged but excluded from field mismatch count.
- Only columns with matching headers are compared.
- Date/number formats are preserved per matching column using the format
  of each column's first data cell.
- Empty data rows are skipped.
- Duplicate or blank Unique IDs stop the run with an explanation.
- Source spreadsheets are read-only when opened by the tool.

TROUBLESHOOTING
- 'Programmatic access to VBProject is not trusted': enable the checkbox
  noted in Installation step 3, and restart Excel before installing.
- 'Macros blocked': ask your IT admin to approve the tool; consider using
  a trusted add-ins location or signing the add-in.
- Ribbon tab missing: restart Excel, then check File > Options > Add-ins >
  Manage: Excel Add-ins > Go and make sure 'DM Compare Tool' is checked.
- Add-in installs but button doesn't work: inspect the VBA code and check
  whether the company's Office macro policy disables VBA entirely.
- Windows PowerShell is used intentionally; Mac Excel cannot run the setup.

This package contains SOURCE and a builder, not a precompiled .xlam.
A real Excel VBA binary add-in is constructed on your Windows PC by the
included setup script using your installed desktop Excel.
