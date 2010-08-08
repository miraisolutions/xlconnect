# Workbook class definition with initialization method
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


setClass("workbook", representation(filename = "character", jobj = "jobjRef"))

setMethod("initialize", "workbook", function(.Object, filename) {
	.Object@filename <- filename
	.Object@jobj <- new(J("com.miraisolutions.xlconnect.integration.r.RWorkbookWrapper"), filename)
	.Object
})
