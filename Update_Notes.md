# Update Notes 

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

