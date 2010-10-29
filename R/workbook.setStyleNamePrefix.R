# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setStyleNamePrefix",
		function(object, prefix) standardGeneric("setStyleNamePrefix"))

setMethod("setStyleNamePrefix", 
		signature(object = "workbook", prefix = "character"), 
		function(object, prefix) {
			jTryCatch(object@jobj$setStyleNamePrefix(prefix))
		}
)
