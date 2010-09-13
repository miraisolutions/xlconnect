# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("existsSheet",
	function(object, name) standardGeneric("existsSheet"))

setMethod("existsSheet", 
		signature(object = "workbook", name = "character"), 
		function(object, name) {
			jTryCatch(object@jobj$existsSheet(name))
		}
)
