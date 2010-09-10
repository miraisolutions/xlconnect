# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.saveWorkbook <- function() {
	
	# Create workbooks
	file.xls <- "resources/testWorkbookSaveWorkbook.xls"
	file.xlsx <- "resources/testWorkbookSaveWorkbook.xlsx"
	wb.xls <- openWorkbook(file.xls, create = TRUE)
	wb.xlsx <- openWorkbook(file.xlsx, create = TRUE)
	
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
