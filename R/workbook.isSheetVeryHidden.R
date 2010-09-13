# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("isSheetVeryHidden",
	function(object, sheet) standardGeneric("isSheetVeryHidden"))

setMethod("isSheetVeryHidden", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet) {
			# Note: Java indices are 0-based
			jTryCatch(object@jobj$isSheetVeryHidden(as.integer(sheet - 1)))
		}
)

setMethod("isSheetVeryHidden", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet) {
			jTryCatch(object@jobj$isSheetVeryHidden(sheet))
		}
)
