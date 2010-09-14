# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.with.workbook <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWithWorkbook.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWithWorkbook.xlsx"), create = FALSE)
	
	# Check if named regions can be correctly referenced (*.xls)
	with(wb.xls, {
		checkTrue(all(dim(AA) == c(8, 3)))
		checkTrue(all(dim(BB) == c(5, 5)))
	}, header = FALSE)

	# Check if named regions can be correctly referenced (*.xlsx)
	with(wb.xlsx, {
		checkTrue(all(dim(AA) == c(8, 3)))
		checkTrue(all(dim(BB) == c(5, 5)))
	}, header = FALSE)
}
