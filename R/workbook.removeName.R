# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("removeName",
	function(object, name) standardGeneric("removeName"))

setMethod("removeName", 
		signature(object = "workbook", name = "character"), 
		function(object, name) {
			jTryCatch(object@jobj$removeName(name))
		}
)
