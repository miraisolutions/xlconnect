# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("readWorksheet",
	function(object, sheet, startRow, startCol, endRow, endCol, header) standardGeneric("readWorksheet"))

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "missing"), 
		function(object, sheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(object@jobj$readWorksheet(as.integer(sheet - 1), TRUE))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "missing"), 
		function(object, sheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(object@jobj$readWorksheet(sheet, TRUE))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "logical"), 
		function(object, sheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(object@jobj$readWorksheet(as.integer(sheet - 1), header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "logical"), 
		function(object, sheet, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(object@jobj$readWorksheet(sheet, header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "logical"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(object@jobj$readWorksheet(as.integer(sheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1), header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "logical"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(object@jobj$readWorksheet(sheet, as.integer(startRow - 1), as.integer(startCol - 1), 
					as.integer(endRow - 1), as.integer(endCol - 1), header))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "missing"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			# note that Java indices are 0-based
			dataFrame <- jTryCatch(object@jobj$readWorksheet(as.integer(sheet - 1), as.integer(startRow - 1), 
							as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1), TRUE))
			dataframeFromJava(dataFrame)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "missing"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			# Read worksheet (returns RDataFrameWrapper Java object reference)
			dataFrame <- jTryCatch(object@jobj$readWorksheet(sheet, as.integer(startRow - 1), as.integer(startCol - 1), 
							as.integer(endRow - 1), as.integer(endCol - 1), TRUE))
			dataframeFromJava(dataFrame)
		}
)
