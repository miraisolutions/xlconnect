# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("readWorksheet")) {
	if(is.function("readWorksheet")) fun <- readWorksheet
	else fun <- function(.Object, worksheet, startRow, startCol, endRow, endCol, header) standardGeneric("readWorksheet")
	setGeneric("readWorksheet", fun)
}

setMethod("readWorksheet", 
		signature(.Object = "workbook", worksheet = "numeric", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "missing"), 
		function(.Object, worksheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(as.integer(worksheet - 1), TRUE))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", worksheet = "character", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "missing"), 
		function(.Object, worksheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(worksheet, TRUE))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", worksheet = "numeric", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "logical"), 
		function(.Object, worksheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(as.integer(worksheet - 1), header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", worksheet = "character", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "logical"), 
		function(.Object, worksheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(worksheet, header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", worksheet = "numeric", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "logical"), 
		function(.Object, worksheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(as.integer(worksheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1), header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", worksheet = "character", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "logical"), 
		function(.Object, worksheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(worksheet, as.integer(startRow - 1), as.integer(startCol - 1), 
					as.integer(endRow - 1), as.integer(endCol - 1), header))
			dataframeFromJava(dataFrame)
		}
)
