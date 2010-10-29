# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.getReferenceFormula <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookReferenceFormula.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReferenceFormula.xlsx"), create = FALSE)
	
	# Check if reference formulas match (*.xls)
	checkTrue(getReferenceFormula(wb.xls, "FirstName") == "Tabelle1!$A$1")
	checkTrue(substring(getReferenceFormula(wb.xls, "SecondName"), 1, 5) == "#REF!")
	
	# Check if reference formulas match (*.xlsx)
	checkTrue(getReferenceFormula(wb.xlsx, "FirstName") == "Tabelle1!$A$1")
	checkTrue(substring(getReferenceFormula(wb.xlsx, "SecondName"), 1, 5) == "#REF!")
	
}
