# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


dataframeToJava <- function(df) {
	# Force data.frame
	if(!is.data.frame(df)) {
		warning("Object is not a data.frame! Trying to continue by converting object.")
		df <- as.data.frame(df)
	}
	
	dFrame <- new(J("com.miraisolutions.xlconnect.integration.r.RDataFrameWrapper"))
	cnames <- colnames(df)
	for(i in seq(along = df)) {
		v <- df[, i]
		
		if(is(v, "numeric")) {
			dFrame$addNumericColumn(cnames[i], .jarray(as.double(v)), .jarray(is.na(v)))
		}
		else if(is(v, "logical")) {
			dFrame$addBooleanColumn(cnames[i], .jarray(v), .jarray(is.na(v)))
		}
		else if(is(v, "character")) {
			dFrame$addStringColumn(cnames[i], .jarray(v), .jarray(is.na(v)))
		}
		else if(is(v, "factor")) {
			v <- as.character(v)
			dFrame$addStringColumn(cnames[i], .jarray(v), .jarray(is.na(v)))
		}
		else if(is(v, "Date") || is(v, "POSIXt")) {
			v <- format(v, format = options("XLConnect.dateTimeFormat")[[1]])
			dFrame$addDateTimeColumn(cnames[i], .jarray(v), .jarray(is.na(v)))
		}
		else
			stop("Unsupported data type (class) detected!")			
	}
	
	dFrame
}
