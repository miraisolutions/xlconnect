# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("getDefinedNames",
	function(object, validOnly) standardGeneric("getDefinedNames"))

setMethod("getDefinedNames", signature(object = "workbook", validOnly = "logical"), function(object, validOnly) {
	jTryCatch(object@jobj$getDefinedNames(validOnly))
})

setMethod("getDefinedNames", signature(object = "workbook", validOnly = "missing"), function(object, validOnly) {
	jTryCatch(object@jobj$getDefinedNames(TRUE))
})
