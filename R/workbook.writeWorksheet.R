# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("writeWorksheet")) {
	if(is.function("writeWorksheet")) fun <- writeWorksheet
	else fun <- function(.Object, data, worksheet, startRow, startCol, 
					header) standardGeneric("writeWorksheet")
	setGeneric("writeWorksheet", fun)
}

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "numeric", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(.Object, data, worksheet, startRow, startCol, header) {
			# note that Java indices are 0-based
			.Object@jobj$writeWorksheet(dataframeToJava(data), as.integer(worksheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1), header)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(.Object, data, worksheet, startRow, startCol, header) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), worksheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), header)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "missing"), 
		function(.Object, data, worksheet, startRow, startCol, header) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), worksheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), TRUE)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(.Object, data, worksheet, startRow, startCol, header) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), as.integer(worksheet - 1), header)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(.Object, data, worksheet, startRow, startCol, header) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), as.integer(worksheet - 1), TRUE)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "character", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(.Object, data, worksheet, startRow, startCol, header) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), worksheet, header)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "character", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(.Object, data, worksheet, startRow, startCol, header) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), worksheet, TRUE)
			invisible()
		}
)
