# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("isSheetVeryHidden")) {
	if(is.function("isSheetVeryHidden")) fun <- isSheetVeryHidden
	else fun <- function(.Object, sheet) standardGeneric("isSheetVeryHidden")
	setGeneric("isSheetVeryHidden", fun)
}

setMethod("isSheetVeryHidden", 
		signature(.Object = "workbook", sheet = "numeric"), 
		function(.Object, sheet) {
			# Note: Java indices are 0-based
			jTryCatch(.Object@jobj$isSheetVeryHidden(as.integer(sheet - 1)))
		}
)

setMethod("isSheetVeryHidden", 
		signature(.Object = "workbook", sheet = "character"), 
		function(.Object, sheet) {
			jTryCatch(.Object@jobj$isSheetVeryHidden(sheet))
		}
)
