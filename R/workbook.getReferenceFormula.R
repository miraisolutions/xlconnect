# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("getReferenceFormula",
		function(object, name) standardGeneric("getReferenceFormula"))

setMethod("getReferenceFormula", 
		signature(object = "workbook", name = "character"), 
		function(object, name) {
			jTryCatch(object@jobj$getReferenceFormula(name))
		}
)
