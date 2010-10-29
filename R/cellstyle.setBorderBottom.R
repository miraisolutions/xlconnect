# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setBorderBottom",
		function(object, type) standardGeneric("setBorderBottom"))

setMethod("setBorderBottom", 
		signature(object = "cellstyle", type = "numeric"), 
		function(object, type) {
			jTryCatch(object@jobj$setBorderBottom(as.integer(type)))
		}
)
