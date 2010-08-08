# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("removeSheet")) {
	if(is.function("removeSheet")) fun <- getSheets
	else fun <- function(.Object) standardGeneric("removeSheet")
	setGeneric("removeSheet", fun)
}

setMethod("removeSheet", "workbook", function(.Object) {
	.Object@jobj$removeSheet()
})
