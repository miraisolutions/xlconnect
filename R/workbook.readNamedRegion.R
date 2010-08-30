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
		dataFrame <- .Object@jobj$readNamedRegion(name, header)
		dataframeFromJava(dataFrame)
	})
