# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("writeNamedRegion")) {
	if(is.function("writeNamedRegion")) fun <- writeNamedRegion
	else fun <- function(.Object, data, name, location, overwrite) standardGeneric("writeNamedRegion")
	setGeneric("writeNamedRegion", fun)
}

setMethod("writeNamedRegion", 
	signature(.Object = "workbook", data = "ANY", name = "character", location = "character", overwrite = "logical"), 
	function(.Object, data, name, location, overwrite) {
		.Object@jobj$writeNamedRegion(dataframeToJava(data), name, location, overwrite)
		invisible()
	}
)

setMethod("writeNamedRegion", 
	signature(.Object = "workbook", data = "ANY", name = "character", location = "missing", overwrite = "missing"), 
	function(.Object, data, name, location, overwrite) {
		.Object@jobj$writeNamedRegion(dataframeToJava(data), name)
		invisible()
	}
)
