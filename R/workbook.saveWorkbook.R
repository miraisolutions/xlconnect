# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


if(!isGeneric("saveWorkbook")) {
	if(is.function("saveWorkbook")) fun <- saveWorkbook
	else fun <- function(.Object) standardGeneric("saveWorkbook")
	setGeneric("saveWorkbook", fun)
}

setMethod("saveWorkbook", signature(.Object = "workbook"), function(.Object) {
	.Object@jobj$save()
})
