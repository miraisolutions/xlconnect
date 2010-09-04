# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("isSheetHidden")) {
	if(is.function("isSheetHidden")) fun <- isSheetHidden
	else fun <- function(.Object, sheet) standardGeneric("isSheetHidden")
	setGeneric("isSheetHidden", fun)
}

setMethod("isSheetHidden", 
		signature(.Object = "workbook", sheet = "numeric"), 
		function(.Object, sheet) {
			# Note: Java indices are 0-based
			jTryCatch(.Object@jobj$isSheetHidden(as.integer(sheet - 1)))
		}
)

setMethod("isSheetHidden", 
		signature(.Object = "workbook", sheet = "character"), 
		function(.Object, sheet) {
			jTryCatch(.Object@jobj$isSheetHidden(sheet))
		}
)
