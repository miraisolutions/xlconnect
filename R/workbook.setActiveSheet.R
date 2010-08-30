# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("setActiveSheet")) {
	if(is.function("setActiveSheet")) fun <- setActiveSheet
	else fun <- function(.Object, sheet) standardGeneric("setActiveSheet")
	setGeneric("setActiveSheet", fun)
}

setMethod("setActiveSheet", 
		signature(.Object = "workbook", sheet = "numeric"), 
		function(.Object, sheet) {
			.Object@jobj$setActiveSheet(as.integer(sheet - 1))
		}
)

setMethod("setActiveSheet", 
		signature(.Object = "workbook", sheet = "character"), 
		function(.Object, sheet) {
			.Object@jobj$setActiveSheet(sheet)
		}
)
