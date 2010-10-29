# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setFillForegroundColor",
		function(object, color) standardGeneric("setFillForegroundColor"))

setMethod("setFillForegroundColor", 
		signature(object = "cellstyle", color = "numeric"), 
		function(object, color) {
			jTryCatch(object@jobj$setFillForeground(as.integer(color)))
		}
)


