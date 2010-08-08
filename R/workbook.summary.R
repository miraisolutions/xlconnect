# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setMethod("summary", "workbook", function(object) {
	cat("Filename: ", object@filename, "\n")
	
	cat("Defined Sheets:\n")
	print(getSheets(object))

	cat("Defined Names:\n")
	print(getDefinedNames(object))
})
