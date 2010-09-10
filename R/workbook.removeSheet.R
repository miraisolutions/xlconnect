# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("removeSheet")) {
	if(is.function("removeSheet")) fun <- removeSheet
	else fun <- function(.Object, sheet) standardGeneric("removeSheet")
	setGeneric("removeSheet", fun)
}

setMethod("removeSheet", 
		signature(.Object = "workbook", sheet = "numeric"), 
		function(.Object, sheet) {
			jTryCatch(.Object@jobj$removeSheet(as.integer(sheet - 1)))
		}
)

setMethod("removeSheet", 
		signature(.Object = "workbook", sheet = "character"), 
		function(.Object, sheet) {
			jTryCatch(.Object@jobj$removeSheet(sheet))
		}
)
