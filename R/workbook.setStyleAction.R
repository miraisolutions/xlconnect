# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setStyleAction",
		function(object, type) standardGeneric("setStyleAction"))

setMethod("setStyleAction", 
		signature(object = "workbook", type = "character"), 
		function(object, type) {
			jTryCatch(object@jobj$setStyleAction(type))
		}
)
