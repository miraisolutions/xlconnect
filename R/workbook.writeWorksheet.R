# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("writeWorksheet")) {
	if(is.function("writeWorksheet")) fun <- writeWorksheet
	else fun <- function(.Object, data, worksheet, startRow, startCol, create) standardGeneric("writeWorksheet")
	setGeneric("writeWorksheet", fun)
}

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "numeric", startRow = "numeric", 
				startCol = "numeric", create = "missing"), 
		function(.Object, data, worksheet, startRow, startCol, create) {
			# note that Java indices are 0-based
			.Object@jobj$writeWorksheet(dataframeToJava(data), as.integer(worksheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "character", startRow = "numeric", 
				startCol = "numeric", create = "logical"), 
		function(.Object, data, worksheet, startRow, startCol, create) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), worksheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), create)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "character", startRow = "numeric", 
				startCol = "numeric", create = "missing"), 
		function(.Object, data, worksheet, startRow, startCol, create) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), worksheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), FALSE)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "numeric", startRow = "missing", 
				startCol = "missing", create = "missing"), 
		function(.Object, data, worksheet, startRow, startCol, create) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), as.integer(worksheet - 1))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "character", startRow = "missing", 
				startCol = "missing", create = "logical"), 
		function(.Object, data, worksheet, startRow, startCol, create) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), worksheet, create)
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(.Object = "workbook", data = "ANY", worksheet = "character", startRow = "missing", 
				startCol = "missing", create = "missing"), 
		function(.Object, data, worksheet, startRow, startCol, create) {
			.Object@jobj$writeWorksheet(dataframeToJava(data), worksheet, FALSE)
			invisible()
		}
)
