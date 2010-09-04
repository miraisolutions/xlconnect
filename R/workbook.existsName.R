# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("existsName")) {
	if(is.function("existsName")) fun <- getSheets
	else fun <- function(.Object, name) standardGeneric("existsName")
	setGeneric("existsName", fun)
}

setMethod("existsName", 
		signature(.Object = "workbook", name = "character"), 
		function(.Object, name) {
			jTryCatch(.Object@jobj$existsName(name))
		}
)
