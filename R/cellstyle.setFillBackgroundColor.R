# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setFillBackgroundColor",
		function(object, color) standardGeneric("setFillBackgroundColor"))

setMethod("setFillBackgroundColor", 
		signature(object = "cellstyle", color = "numeric"), 
		function(object, color) {
			jTryCatch(object@jobj$setFillBackgroundColor(as.integer(color)))
		}
)
