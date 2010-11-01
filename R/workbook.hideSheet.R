# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################
 
setGeneric("hideSheet",
	function(object, sheet, veryHidden) standardGeneric("hideSheet"))

setMethod("hideSheet", 
		signature(object = "workbook", sheet = "numeric", veryHidden = "logical"), 
		function(object, sheet, veryHidden) {
			object@jobj$hideSheet(as.integer(sheet - 1), veryHidden)
		}
)

setMethod("hideSheet", 
		signature(object = "workbook", sheet = "numeric", veryHidden = "missing"), 
		function(object, sheet, veryHidden) {
			callGeneric(object, sheet, FALSE)
		}
)

setMethod("hideSheet", 
		signature(object = "workbook", sheet = "character", veryHidden = "logical"), 
		function(object, sheet, veryHidden) {
			object@jobj$hideSheet(sheet, veryHidden)
		}
)

setMethod("hideSheet", 
		signature(object = "workbook", sheet = "character", veryHidden = "missing"), 
		function(object, sheet, veryHidden) {
			callGeneric(object, sheet, FALSE)
		}
)
