# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("isSheetHidden",
	function(object, sheet) standardGeneric("isSheetHidden"))

setMethod("isSheetHidden", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet) {
			# Note: Java indices are 0-based
			jTryCatch(object@jobj$isSheetHidden(as.integer(sheet - 1)))
		}
)

setMethod("isSheetHidden", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet) {
			jTryCatch(object@jobj$isSheetHidden(sheet))
		}
)
