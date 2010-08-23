# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

prepareForXLConnect <- function(v) {
	
	if(is(v, "numeric")) {
		dataType <- "Numeric"
	}
	else if(is(v, "logical")) {
		dataType <- "Boolean"
	}
	else if(is(v, "character")) {
		dataType <- "String"
	}
	else if(is(v, "factor")) {
		dataType <- "String"
		v <- as.character(v)
	}
	else if(is(v, "POSIXt")) {
		dataType <- "DateTime"
		v <- format(v, format = options("XLConnect.dateTimeFormat")[[1]])
	}
	else if(is(v, "Date")) {
		dataType <- "DateTime"
		# Convert Date to POSIXlt before formatting as character
		v <- format(as.POSIXlt(v), format = options("XLConnect.dateTimeFormat")[[1]])
	}
	else
		stop("Unsupported data type (class) detected!")
	
	list(data = v, dataType = dataType)
}
