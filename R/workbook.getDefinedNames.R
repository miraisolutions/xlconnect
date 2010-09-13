# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("getDefinedNames",
	function(object) standardGeneric("getDefinedNames"))

setMethod("getDefinedNames", signature(object = "workbook"), function(object) {
	jTryCatch(object@jobj$getDefinedNames())
})
