# Workbook class definition with initialization method
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


setClass("workbook", representation(filename = "character", jobj = "jobjRef"))

setMethod("initialize", 
		"workbook", 
		function(.Object, filename, create) {
			.Object@filename <- filename
			.Object@jobj <- jTryCatch(new(J("com.miraisolutions.xlconnect.integration.r.RWorkbookWrapper"), filename, create))
			if(is.jnull(.Object@jobj))
				stop("Could not create workbook instance! Got null reference from Java.")
			.Object
		}
)
