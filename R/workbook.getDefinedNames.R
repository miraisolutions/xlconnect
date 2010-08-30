# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("getDefinedNames")) {
	if(is.function("getDefinedNames")) fun <- getDefinedNames
	else fun <- function(.Object) standardGeneric("getDefinedNames")
	setGeneric("getDefinedNames", fun)
}

setMethod("getDefinedNames", signature(.Object = "workbook"), function(.Object) {
	.Object@jobj$getDefinedNames()
})
