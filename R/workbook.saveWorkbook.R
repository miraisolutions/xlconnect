# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("saveWorkbook",
	function(object) standardGeneric("saveWorkbook"))

setMethod("saveWorkbook", signature(object = "workbook"), function(object) {
	jTryCatch(object@jobj$save())
})
