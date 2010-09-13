# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("writeWorksheet")) {
	if(is.function("writeWorksheet")) fun <- writeWorksheet
	else fun <- function(.Object, data, sheet, startRow, startCol, 
					header) standardGeneric("writeWorksheet")
	setGeneric("writeWorksheet", fun)
}

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", sheet = "numeric", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(.Object, data, sheet, startRow, startCol, header) {
			# note that Java indices are 0-based
			jTryCatch(.Object@jobj$writeWorksheet(dataframeToJava(data), as.integer(sheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", sheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(.Object, data, sheet, startRow, startCol, header) {
			jTryCatch(.Object@jobj$writeWorksheet(dataframeToJava(data), sheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", sheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "missing"), 
		function(.Object, data, sheet, startRow, startCol, header) {
			jTryCatch(.Object@jobj$writeWorksheet(dataframeToJava(data), sheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), TRUE))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(.Object, data, sheet, startRow, startCol, header) {
			jTryCatch(.Object@jobj$writeWorksheet(dataframeToJava(data), as.integer(sheet - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", sheet = "character", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(.Object, data, sheet, startRow, startCol, header) {
			jTryCatch(.Object@jobj$writeWorksheet(dataframeToJava(data), sheet, header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(.Object, data, sheet, startRow, startCol, header) {
			jTryCatch(.Object@jobj$writeWorksheet(dataframeToJava(data), as.integer(sheet - 1), TRUE))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", sheet = "character", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(.Object, data, sheet, startRow, startCol, header) {
			jTryCatch(.Object@jobj$writeWorksheet(dataframeToJava(data), sheet, TRUE))
			invisible()
		}
)
