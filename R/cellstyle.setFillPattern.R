# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setFillPattern",
		function(object, fill) standardGeneric("setFillPattern"))

setMethod("setFillPattern", 
		signature(object = "cellstyle", fill = "numeric"), 
		function(object, fill) {
			jTryCatch(object@jobj$setFillPattern(as.integer(fill)))
		}
)
