# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("addColumn")) {
	if(is.function("addColumn")) fun <- getSheets
	else fun <- function(.Object, name, dataType, data) standardGeneric("addColumn")
	setGeneric("addColumn", fun)
}

setMethod("addColumn", 
		signature(.Object = "dframe", name = "character", dataType = "character", data = "ANY"), 
		function(.Object, name, dataType, data) {
	 
	if(dataType == "Numeric") {
		.Object@jobj$addNumericColumn(name, .jarray(data), .jarray(is.na(data)))
	}
	else if(dataType == "Boolean") {
		.Object@jobj$addBooleanColumn(name, .jarray(data), .jarray(is.na(data)))
	}
	else if(dataType == "String") {
		.Object@jobj$addStringColumn(name, .jarray(data), .jarray(is.na(data)))
	}
	else if(dataType == "DateTime") {
		.Object@jobj$addDateTimeColumn(name, .jarray(data), .jarray(is.na(data)))
	}
	else
		stop("Unsupported data type!")
	
	invisible()
})
