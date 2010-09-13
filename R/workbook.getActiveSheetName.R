# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("getActiveSheetName")) {
	if(is.function("getActiveSheetName")) fun <- getActiveSheetName
	else fun <- function(.Object) standardGeneric("getActiveSheetName")
	setGeneric("getActiveSheetName", fun)
}

setMethod("getActiveSheetName", 
		signature(.Object = "workbook"), 
		function(.Object) {
			sheet <- jTryCatch(.Object@jobj$getActiveSheetName())
			ifelse(is.null(sheet), NA, sheet)
		}
)
