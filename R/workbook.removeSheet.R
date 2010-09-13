# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################
 
setGeneric("removeSheet",
	function(object, sheet) standardGeneric("removeSheet"))

setMethod("removeSheet", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet) {
			jTryCatch(object@jobj$removeSheet(as.integer(sheet - 1)))
		}
)

setMethod("removeSheet", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet) {
			jTryCatch(object@jobj$removeSheet(sheet))
		}
)
