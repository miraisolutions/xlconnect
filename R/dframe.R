# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setClass("dframe", representation(jobj = "jobjRef"))

setMethod("initialize", "dframe", function(.Object) {
	.Object@jobj <- new(J("com.miraisolutions.xlconnect.integration.r.RDataFrameWrapper"))
	.Object
})
