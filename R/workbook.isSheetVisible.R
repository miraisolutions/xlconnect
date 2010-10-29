# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("isSheetVisible",
		function(object, sheet) standardGeneric("isSheetVisible"))

setMethod("isSheetVisible", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet) {
			!isSheetHidden(object, sheet) && !isSheetVeryHidden(object, sheet)
		}
)

setMethod("isSheetVisible", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet) {
			!isSheetHidden(object, sheet) && !isSheetVeryHidden(object, sheet)
		}
)
