# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setMethod("print", signature(x = "workbook"), function(x, ...) {
	x@filename
})
