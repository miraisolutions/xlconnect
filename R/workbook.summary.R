# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setMethod("summary", signature(object = "workbook"), function(object) {
	cat("*** XLConnect Workbook Summary ***\n")
	cat("> Filename: '", object@filename, "'\n", sep = "")
	
	nice <- function(x) {
		if(length(x) > 0) x
		else "<NONE>"
	}
	
	sheets <- getSheets(object)
	
	cat("> Sheets (all):\n")
	cat(nice(sheets), sep = ", ", fill = TRUE)
	
	cat("> Hidden Sheets:\n")
	idx <- sapply(sheets, function(s) isSheetHidden(object, s))
	cat(nice(sheets[idx]), sep = ", ", fill = TRUE)
	
	cat("> Very Hidden Sheets:\n")
	idx <- sapply(sheets, function(s) isSheetVeryHidden(object, s))
	cat(nice(sheets[idx]), sep = ", ", fill = TRUE)
	
	cat("> Names:\n")
	cat(nice(getDefinedNames(object)), sep = ", ", fill = TRUE)
	
	cat("> Active Sheet: ", nice(getActiveSheetName(object)), "\n")
})
