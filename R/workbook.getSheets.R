# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("getSheets")) {
	if(is.function("getSheets")) fun <- getSheets
	else fun <- function(.Object) standardGeneric("getSheets")
	setGeneric("getSheets", fun)
}

setMethod("getSheets", "workbook", function(.Object) {
	.Object@jobj$getSheets()
})
