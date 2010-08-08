# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("createSheet")) {
	if(is.function("createSheet")) fun <- getSheets
	else fun <- function(.Object, name) standardGeneric("createSheet")
	setGeneric("createSheet", fun)
}

setMethod("createSheet", "workbook", function(.Object, name) {
	.Object@jobj$createSheet(name)
})
