# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("getActiveSheetName",
	function(object) standardGeneric("getActiveSheetName"))

setMethod("getActiveSheetName", 
		signature(object = "workbook"), 
		function(object) {
			sheet <- jTryCatch(object@jobj$getActiveSheetName())
			ifelse(is.null(sheet), NA, sheet)
		}
)
