# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("getActiveSheetIndex")) {
	if(is.function("getActiveSheetIndex")) fun <- getActiveSheetIndex
	else fun <- function(.Object) standardGeneric("getActiveSheetIndex")
	setGeneric("getActiveSheetIndex", fun)
}

setMethod("getActiveSheetIndex", 
		signature(.Object = "workbook"), 
		function(.Object) {
			# Note: Java has 0-based indices
			idx <- as.integer(jTryCatch(.Object@jobj$getActiveSheetIndex()) + 1)
			ifelse(idx > 0, idx, NA)
		}
)
