# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("setBorder",
		function(object, side, type, color) standardGeneric("setBorder"))

setMethod("setBorder", 
		signature(object = "cellstyle", side = "character", type = "numeric", 
				color = "numeric"), 
		function(object, side, type, color) {
			side <- tolower(side)
			if("all" %in% side)
				side <- c("bottom", "left", "right", "top")
			
			type <- as.integer(rep(type, length = length(side)))
			color <- as.integer(rep(color, length = length(side)))
			
			jTryCatch(object@jobj$setBorder(.jarray(side), .jarray(type), .jarray(color)))
		}
)
