# Update Notes 

### 12-7-2021
 - Removed the following modules:
	 - Reports
	 - PotatoPrint
	 - lotta subs from SheetGenerators
	 - ClearAndReturn
 - Removed the TimeTableQuery and sheet
 - Reworked refreshNfilterModule. It now refreshes only prod_mergeSheets, im.blendQty.onHand, and IssueSheetTable
 - Added WorkbookOpen(), which does the following when ISSUESHEET_GEN.xlsb is opened:
	 1. open ProdScheduleCopy.xlsb
	 2. open NOT_Blending_Issue_Sheet.xlsb

