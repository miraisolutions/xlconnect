# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setDataFormat",
		function(object, format) standardGeneric("setDataFormat"))

setMethod("setDataFormat", 
		signature(object = "cellstyle", format = "character"), 
		function(object, format) {
			jTryCatch(object@jobj$setDataFormat(format))
		}
)
