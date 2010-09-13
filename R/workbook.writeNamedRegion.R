# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("writeNamedRegion",
	function(object, data, name, header) standardGeneric("writeNamedRegion"))

setMethod("writeNamedRegion", 
	signature(object = "workbook", data = "ANY", name = "character", header = "logical"), 
	function(object, data, name, header) {
		jTryCatch(object@jobj$writeNamedRegion(dataframeToJava(data), name, header))
		invisible()
	}
)

setMethod("writeNamedRegion", 
		signature(object = "workbook", data = "ANY", name = "character", header = "missing"), 
		function(object, data, name, header) {
			jTryCatch(object@jobj$writeNamedRegion(dataframeToJava(data), name, TRUE))
			invisible()
		}
)
