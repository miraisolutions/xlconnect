# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("writeWorksheet",
	function(object, data, sheet, startRow, startCol, header) standardGeneric("writeWorksheet"))

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			# note that Java indices are 0-based
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), as.integer(sheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), sheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), sheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), TRUE))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), as.integer(sheet - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), sheet, header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), as.integer(sheet - 1), TRUE))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), sheet, TRUE))
			invisible()
		}
)
