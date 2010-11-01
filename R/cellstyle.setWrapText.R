# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setWrapText",
		function(object, wrap) standardGeneric("setWrapText"))

setMethod("setWrapText", 
		signature(object = "cellstyle", wrap = "logical"), 
		function(object, wrap) {
			jTryCatch(object@jobj$setWrapText(wrap))
		}
)
