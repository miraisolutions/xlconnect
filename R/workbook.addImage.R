# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("addImage")) {
	if(is.function("addImage")) fun <- addImage
	else fun <- function(.Object, filename, name, originalSize,
				location, overwrite) standardGeneric("addImage")
	setGeneric("addImage", fun)
}

setMethod("addImage", 
		signature(.Object = "workbook", filename = "character", name = "character", 
				originalSize = "logical", location = "character", overwrite = "logical"), 
		function(.Object, filename, name, originalSize, location, overwrite) {
			jTryCatch(.Object@jobj$addImage(filename, originalSize, name, location, overwrite))
		}
)

setMethod("addImage", 
		signature(.Object = "workbook", filename = "character", name = "character", 
				originalSize = "logical", location = "character", overwrite = "missing"), 
		function(.Object, filename, name, originalSize, location, overwrite) {
			jTryCatch(.Object@jobj$addImage(filename, originalSize, name, location, FALSE))
		}
)

setMethod("addImage", 
		signature(.Object = "workbook", filename = "character", name = "character", 
				originalSize = "logical", location = "missing", overwrite = "missing"), 
		function(.Object, filename, name, originalSize, location, overwrite) {
			jTryCatch(.Object@jobj$addImage(filename, originalSize, name))
		}
)

setMethod("addImage", 
		signature(.Object = "workbook", filename = "character", name = "character", 
				originalSize = "missing", location = "missing", overwrite = "missing"), 
		function(.Object, filename, name, originalSize, location, overwrite) {
			jTryCatch(.Object@jobj$addImage(filename, FALSE, name))
		}
)
