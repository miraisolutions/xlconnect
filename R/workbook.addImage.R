# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("addImage", 
	function(object, filename, name, originalSize) standardGeneric("addImage"))

setMethod("addImage", 
		signature(object = "workbook", filename = "character", name = "character", 
				originalSize = "logical"), 
		function(object, filename, name, originalSize) {
			jTryCatch(object@jobj$addImage(filename, name, originalSize))
		}
)

setMethod("addImage", 
		signature(object = "workbook", filename = "character", name = "character", 
				originalSize = "missing"), 
		function(object, filename, name, originalSize) {
			jTryCatch(object@jobj$addImage(filename, name, FALSE))
		}
)
