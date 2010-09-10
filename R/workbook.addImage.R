# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("addImage")) {
	if(is.function("addImage")) fun <- addImage
	else fun <- function(.Object, filename, name, originalSize) standardGeneric("addImage")
	setGeneric("addImage", fun)
}

setMethod("addImage", 
		signature(.Object = "workbook", filename = "character", name = "character", 
				originalSize = "logical"), 
		function(.Object, filename, name, originalSize) {
			jTryCatch(.Object@jobj$addImage(filename, name, originalSize))
		}
)

setMethod("addImage", 
		signature(.Object = "workbook", filename = "character", name = "character", 
				originalSize = "missing"), 
		function(.Object, filename, name, originalSize) {
			jTryCatch(.Object@jobj$addImage(filename, name, FALSE))
		}
)
