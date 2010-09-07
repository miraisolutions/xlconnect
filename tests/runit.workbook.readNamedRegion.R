# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.readNamedRegion <- function() {
	
	# Create workbooks
	wb.xls <- openWorkbook("resources/testWorkbookReadNamedRegion.xls", create = FALSE)
	wb.xlsx <- openWorkbook("resources/testWorkbookReadNamedRegion.xlsx", create = FALSE)
	
	checkDf <- data.frame(
			"NumericColumn" = c(-23.63, NA, NA, 5.8, 3),
			"StringColumn" = c("Hello", NA, NA, NA, "World"),
			"BooleanColumn" = c(TRUE, FALSE, FALSE, NA, NA),
			"DateTimeColumn" = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07"), tz = "UTC"),
			stringsAsFactors = F
	)
	
	# Check that the read named region equals the defined data.frame (*.xls)
	res <- readNamedRegion(wb.xls, "Test", header = TRUE)
	checkEquals(res, checkDf)
	
	# Check that the read named region equals the defined data.frame (*.xlsx)
	res <- readNamedRegion(wb.xlsx, "Test", header = TRUE)
	checkEquals(res, checkDf)
	
	# Check that attempting to read a non-existing named region throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "NameThatDoesNotExist"))
	
	# Check that attempting to read a non-existing named region throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "NameThatDoesNotExist"))
}
