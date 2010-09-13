# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("readNamedRegion",
	function(object, name, header) standardGeneric("readNamedRegion"))

setMethod("readNamedRegion", 
	signature(object = "workbook", name = "character", header = "logical"), 
	function(object, name, header) {	
		# Read named region (returns RDataFrameWrapper Java object reference)
		dataFrame <- jTryCatch(object@jobj$readNamedRegion(name, header))
		dataframeFromJava(dataFrame)
	}
)

setMethod("readNamedRegion", 
		signature(object = "workbook", name = "character", header = "missing"), 
		function(object, name, header) {	
			# Read named region (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(object@jobj$readNamedRegion(name, TRUE))
			dataframeFromJava(dataFrame)
		}
)
