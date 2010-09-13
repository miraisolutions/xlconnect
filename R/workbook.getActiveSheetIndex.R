# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("getActiveSheetIndex",
	function(object) standardGeneric("getActiveSheetIndex"))

setMethod("getActiveSheetIndex", 
		signature(object = "workbook"), 
		function(object) {
			# Note: Java has 0-based indices
			idx <- as.integer(jTryCatch(object@jobj$getActiveSheetIndex()) + 1)
			ifelse(idx > 0, idx, NA)
		}
)
