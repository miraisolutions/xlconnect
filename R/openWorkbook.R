# Helper function for opening workbooks
#
# This is to be able to write openWorkbook(...) rather than new("workbook", filename = ...)
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

openWorkbook <- function(filename, create = FALSE) {
	new("workbook", filename = filename, create = create)
}
