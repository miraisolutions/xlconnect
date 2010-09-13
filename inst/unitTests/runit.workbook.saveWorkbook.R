# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.saveWorkbook <- function() {
	
	# Create workbooks
	file.xls <- rsrc("resources/testWorkbookSaveWorkbook.xls")
	file.xlsx <- rsrc("resources/testWorkbookSaveWorkbook.xlsx")
	file.remove(file.xls)
	file.remove(file.xlsx)
	wb.xls <- loadWorkbook(file.xls, create = TRUE)
	wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
	
	# Files don't exist yet
	checkTrue(!file.exists(file.xls))
	checkTrue(!file.exists(file.xlsx))
	
	saveWorkbook(wb.xls)
	saveWorkbook(wb.xlsx)
	
	# Check that file exists after saving (*.xls)
	checkTrue(file.exists(file.xls))
	
	# Check that file exists after saving (*.xlsx)
	checkTrue(file.exists(file.xlsx))
}
