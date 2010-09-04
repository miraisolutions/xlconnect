# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.existsSheet <- function() {
	
	# Create workbooks
	wb.xls <- openWorkbook("resources/testWorkbookExistsNameAndSheet.xls", create = FALSE)
	wb.xlsx <- openWorkbook("resources/testWorkbookExistsNameAndSheet.xlsx", create = FALSE)
	
	# Check that the following sheets exists (*.xls)
	checkTrue(existsSheet(wb.xls, "AAA"))
	checkTrue(existsSheet(wb.xls, "BBB"))
	checkTrue(existsSheet(wb.xls, "CCC"))
	
	# Check that the following do NOT exists (*.xls)
	checkTrue(!existsSheet(wb.xls, "DDD"))
	checkTrue(!existsSheet(wb.xls, "'illegal name"))
	checkTrue(!existsSheet(wb.xls, "%&$$-^~@afk20 235-ääaü"))
	
	# Check that the following names exists (*.xlsx)
	checkTrue(existsSheet(wb.xlsx, "AAA"))
	checkTrue(existsSheet(wb.xlsx, "BBB"))
	checkTrue(existsSheet(wb.xlsx, "CCC"))
	
	# Check that the following do NOT exists (*.xlsx)
	checkTrue(!existsSheet(wb.xlsx, "DDD"))
	checkTrue(!existsSheet(wb.xlsx, "'illegal name"))
	checkTrue(!existsSheet(wb.xlsx, "%&$$-^~@afk20 235-ääaü"))
	
}
