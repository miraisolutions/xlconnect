# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("existsName",
	function(object, name) standardGeneric("existsName"))

setMethod("existsName", 
		signature(object = "workbook", name = "character"), 
		function(object, name) {
			jTryCatch(object@jobj$existsName(name))
		}
)
