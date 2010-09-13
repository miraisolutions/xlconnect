# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("unhideSheet",
	function(object, sheet) standardGeneric("unhideSheet"))

setMethod("unhideSheet", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet) {
			jTryCatch(object@jobj$unhideSheet(as.integer(sheet - 1)))
		}
)

setMethod("unhideSheet", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet) {
			jTryCatch(object@jobj$unhideSheet(sheet))
		}
)
