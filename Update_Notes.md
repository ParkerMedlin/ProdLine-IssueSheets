# Update Notes 

### 3-15-2022 
 - changed logic so it no longer deletes sheet 1 in the destination template
 - updated path so that createIssueWorkbooks now has its own destination template, resolving issues with hx blends

### 1-25-2022 
 - Resolved formatting problem with issue sheet; it was skipping too many lines before pasting the next "page" into the document. Just changed increment for `pasteRowXYZLine` variables from 40 to 39 

### 1-18-2022
 - cleared out the bom.blend, timetable, and checkoutcounts queries
 - change name of bom.master to bom.Source and made it more specific, like the one in Blend Schedule 
 - updated formulae on blendData so they are now keyed to bom.Source
 - updated refresh macro to just RefreshAll

### 1-13-2022 
 - FINALLY GOT THE THING TO PRINT MY ISSUE SHEETS TO THE FRONT OFFICE PRINTER WITH STAPLES. Added the macro for this. 

### 1-3-2022
 - Big ole update drastically decreasing the amount of code needed to generate the issue sheets 
 - Eventual goal is to completely automate the staple printing but for now printing will be manual. 
 - createIssueWorkbooks now just copies each set of issue sheets into a different tab and it ends with opening the new workbook

### 12-22-2021
 - Started playing with the option of auto staple printing but it's going to take a lot of work because printer will not staple different worksheets together. 

### 12-21-2021 
 - Removed some unnecessary variables from the CreateIssueSheetWorkbooks subroutine

### 12-16-2021
 - simplified and cleaned up the issue sheet generating loop and its sub functions. removed a couple functions that didn't need to be separate. 
	 - moved the batch number SQL query from the issue sheet template workbook over to the ISSUESHEET_GEN workbook to avoid potential issue of batch formulas failing to calculate before the code executes.

### 12-8-2021
 - resolved the "Range method of object _Global has failed" issue. Problem was that I was trying to select a table row using the old table name (I updated the table name yesterday to match naming convention of all other tables).
 - a-dded `if` statement in the ThisWorkbook module of ISSUESHEET_GEN workbook to make sure Open_Workbook() isn't looking for ProdScheduleCopy.xlsb when it's not there. -
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

