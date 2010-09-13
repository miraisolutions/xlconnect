# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("readWorksheet")) {
	if(is.function("readWorksheet")) fun <- readWorksheet
	else fun <- function(.Object, sheet, startRow, startCol, endRow, endCol, header) standardGeneric("readWorksheet")
	setGeneric("readWorksheet", fun)
}

setMethod("readWorksheet", 
		signature(.Object = "workbook", sheet = "numeric", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "missing"), 
		function(.Object, sheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(as.integer(sheet - 1), TRUE))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", sheet = "character", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "missing"), 
		function(.Object, sheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(sheet, TRUE))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", sheet = "numeric", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "logical"), 
		function(.Object, sheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(as.integer(sheet - 1), header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", sheet = "character", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "logical"), 
		function(.Object, sheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(sheet, header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", sheet = "numeric", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "logical"), 
		function(.Object, sheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(as.integer(sheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1), header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", sheet = "character", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "logical"), 
		function(.Object, sheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(sheet, as.integer(startRow - 1), as.integer(startCol - 1), 
					as.integer(endRow - 1), as.integer(endCol - 1), header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", sheet = "numeric", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "missing"), 
		function(.Object, sheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(as.integer(sheet - 1), as.integer(startRow - 1), 
							as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1), TRUE))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(.Object = "workbook", sheet = "character", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "missing"), 
		function(.Object, sheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(.Object@jobj$readWorksheet(sheet, as.integer(startRow - 1), as.integer(startCol - 1), 
							as.integer(endRow - 1), as.integer(endCol - 1), TRUE))
			dataframeFromJava(dataFrame)
		}
)
