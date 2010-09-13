# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.existsName <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xls", create = FALSE)
	wb.xlsx <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xlsx", create = FALSE)
	
	# Check that the following names exists (*.xls)
	checkTrue(existsName(wb.xls, "AA"))
	checkTrue(existsName(wb.xls, "BB"))
	checkTrue(existsName(wb.xls, "CC"))
	
	# Check that the following do NOT exists (*.xls)
	checkTrue(!existsName(wb.xls, "DD"))
	checkTrue(!existsName(wb.xls, "'illegal name"))
	checkTrue(!existsName(wb.xls, "%&$$-^~@afk20 235-ääaü"))
	
	# Check that the following names exists (*.xlsx)
	checkTrue(existsName(wb.xlsx, "AA"))
	checkTrue(existsName(wb.xlsx, "BB"))
	checkTrue(existsName(wb.xlsx, "CC"))
	
	# Check that the following do NOT exists (*.xlsx)
	checkTrue(!existsName(wb.xlsx, "DD"))
	checkTrue(!existsName(wb.xlsx, "'illegal name"))
	checkTrue(!existsName(wb.xlsx, "%&$$-^~@afk20 235-ääaü"))
}
