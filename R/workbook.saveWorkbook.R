# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


if(!isGeneric("saveWorkbook")) {
	if(is.function("saveWorkbook")) fun <- save
	else fun <- function(.Object) standardGeneric("saveWorkbook")
	setGeneric("saveWorkbook", fun)
}

setMethod("saveWorkbook", "workbook", function(.Object) {
	.Object@jobj$save()
})
