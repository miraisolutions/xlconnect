# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("writeNamedRegion")) {
	if(is.function("writeNamedRegion")) fun <- writeNamedRegion
	else fun <- function(.Object, data, name, header) standardGeneric("writeNamedRegion")
	setGeneric("writeNamedRegion", fun)
}

setMethod("writeNamedRegion", 
	signature(.Object = "workbook", data = "ANY", name = "character", header = "logical"), 
	function(.Object, data, name, header) {
		.Object@jobj$writeNamedRegion(dataframeToJava(data), name, header)
		invisible()
	}
)

setMethod("writeNamedRegion", 
		signature(.Object = "workbook", data = "ANY", name = "character", header = "missing"), 
		function(.Object, data, name, header) {
			.Object@jobj$writeNamedRegion(dataframeToJava(data), name, TRUE)
			invisible()
		}
)
