# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("getSheets",
	function(object) standardGeneric("getSheets"))

setMethod("getSheets", signature(object = "workbook"), function(object) {
	jTryCatch(object@jobj$getSheets())
})
