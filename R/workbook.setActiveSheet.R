# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################
 
setGeneric("setActiveSheet",
	function(object, sheet) standardGeneric("setActiveSheet"))

setMethod("setActiveSheet", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet) {
			jTryCatch(object@jobj$setActiveSheet(as.integer(sheet - 1)))
		}
)

setMethod("setActiveSheet", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet) {
			jTryCatch(object@jobj$setActiveSheet(sheet))
		}
)
