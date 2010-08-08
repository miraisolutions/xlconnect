# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("getDefinedNames")) {
	if(is.function("getDefinedNames")) fun <- getSheets
	else fun <- function(.Object) standardGeneric("getDefinedNames")
	setGeneric("getDefinedNames", fun)
}

setMethod("getDefinedNames", "workbook", function(.Object) {
	.Object@jobj$getDefinedNames()
})
