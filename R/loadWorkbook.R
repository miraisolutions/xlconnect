# Helper function for opening workbooks
#
# This is to be able to write loadWorkbook(...) rather than new("workbook", filename = ...)
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

loadWorkbook <- function(filename, create = FALSE) {
	jTryCatch(new("workbook", filename = filename, create = create))
}
