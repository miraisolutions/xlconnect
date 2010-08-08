# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setMethod("print", "workbook", function(x) {
	x@filename
})
