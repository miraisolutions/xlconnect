# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("readNamedRegion")) {
	if(is.function("readNamedRegion")) fun <- readNamedRegion
	else fun <- function(.Object, name, header) standardGeneric("readNamedRegion")
	setGeneric("readNamedRegion", fun)
}

setMethod("readNamedRegion", 
	signature(.Object = "workbook", name = "character", header = "logical"), 
	function(.Object, name, header) {	
		# Read named region (returns RDataFrameWrapper Java object reference)
		dataFrame <- jTryCatch(.Object@jobj$readNamedRegion(name, header))
		dataframeFromJava(dataFrame)
	}
)

setMethod("readNamedRegion", 
		signature(.Object = "workbook", name = "character", header = "missing"), 
		function(.Object, name, header) {	
			# Read named region (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(.Object@jobj$readNamedRegion(name, TRUE))
			dataframeFromJava(dataFrame)
		}
)
