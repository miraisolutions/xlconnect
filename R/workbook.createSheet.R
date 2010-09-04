# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("createSheet")) {
	if(is.function("createSheet")) fun <- createSheet
	else fun <- function(.Object, name) standardGeneric("createSheet")
	setGeneric("createSheet", fun)
}

setMethod("createSheet", signature(.Object = "workbook", name = "character"), function(.Object, name) {
	jTryCatch(.Object@jobj$createSheet(name))
})
