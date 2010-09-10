# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("unhideSheet")) {
	if(is.function("unhideSheet")) fun <- unhideSheet
	else fun <- function(.Object, sheet) standardGeneric("unhideSheet")
	setGeneric("unhideSheet", fun)
}

setMethod("unhideSheet", 
		signature(.Object = "workbook", sheet = "numeric"), 
		function(.Object, sheet) {
			jTryCatch(.Object@jobj$unhideSheet(as.integer(sheet - 1)))
		}
)

setMethod("unhideSheet", 
		signature(.Object = "workbook", sheet = "character"), 
		function(.Object, sheet) {
			jTryCatch(.Object@jobj$unhideSheet(sheet))
		}
)
