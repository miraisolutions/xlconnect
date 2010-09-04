# Display workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setMethod("show", signature(object = "workbook"), function(object) {
	print(object)
})
