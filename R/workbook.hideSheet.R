# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("hideSheet")) {
	if(is.function("hideSheet")) fun <- hideSheet
	else fun <- function(.Object, sheet, veryHidden) standardGeneric("hideSheet")
	setGeneric("hideSheet", fun)
}

setMethod("hideSheet", 
		signature(.Object = "workbook", sheet = "numeric", veryHidden = "logical"), 
		function(.Object, sheet, veryHidden) {
			.Object@jobj$hideSheet(as.integer(sheet - 1), veryHidden)
		}
)

setMethod("hideSheet", 
		signature(.Object = "workbook", sheet = "numeric", veryHidden = "missing"), 
		function(.Object, sheet, veryHidden) {
			.Object@jobj$hideSheet(as.integer(sheet - 1), FALSE)
		}
)

setMethod("hideSheet", 
		signature(.Object = "workbook", sheet = "character", veryHidden = "logical"), 
		function(.Object, sheet, veryHidden) {
			.Object@jobj$hideSheet(sheet, veryHidden)
		}
)

setMethod("hideSheet", 
		signature(.Object = "workbook", sheet = "character", veryHidden = "missing"), 
		function(.Object, sheet, veryHidden) {
			.Object@jobj$hideSheet(sheet, FALSE)
		}
)
