# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("existsSheet")) {
	if(is.function("existsSheet")) fun <- getSheets
	else fun <- function(.Object, name) standardGeneric("existsSheet")
	setGeneric("existsSheet", fun)
}

setMethod("existsSheet", 
		signature(.Object = "workbook", name = "character"), 
		function(.Object, name) {
			jTryCatch(.Object@jobj$existsSheet(name))
		}
)
