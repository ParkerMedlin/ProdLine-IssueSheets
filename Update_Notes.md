# Update Notes 

### 12-22-2021
 - Started playing with the option of auto staple printing but it's going to take a lot of work because printer will not staple different worksheets together. 

### 12-21-2021 
 - Removed some unnecessary variables from the CreateIssueSheetWorkbooks subroutine

### 12-16-2021
 - simplified and cleaned up the issue sheet generating loop and its sub functions. removed a couple functions that didn't need to be separate. 
	 - moved the batch number SQL query from the issue sheet template workbook over to the ISSUESHEET_GEN workbook to avoid potential issue of batch formulas failing to calculate before the code executes.

### 12-8-2021
 - resolved the "Range method of object _Global has failed" issue. Problem was that I was trying to select a table row using the old table name (I updated the table name yesterday to match naming convention of all other tables).
 - added `if` statement in the ThisWorkbook module of ISSUESHEET_GEN workbook to make sure Open_Workbook() isn't looking for ProdScheduleCopy.xlsb when it's not there. 
 - Removed a couple functions from CreateIssueWorkbooks module.
	 - Sub OpenNOTIssueSht()
	 - Sub OpenWkbkNames() 

### 12-7-2021
 - Removed the following modules:
	 - Reports
	 - PotatoPrint
	 - lotta subs from SheetGenerators
	 - ClearAndReturn
 - Removed the TimeTableQuery and sheet
 - Reworked refreshNfilterModule. It now refreshes only prod_mergeSheets, im.blendQty.onHand, and IssueSheetTable
 - Added WorkbookOpen(), which does the following when ISSUESHEET_GEN.xlsb is opened:
	 1.  open NOT_Blending_Issue_Sheet.xlsb
	 2.  open ProdScheduleCopy.xlsb
	 3.  activate INLINE sheet in prod schedule copy

