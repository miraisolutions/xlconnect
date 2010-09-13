# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("createSheet",
	function(object, name) standardGeneric("createSheet"))

setMethod("createSheet", signature(object = "workbook", name = "character"), function(object, name) {
	jTryCatch(object@jobj$createSheet(name))
})
