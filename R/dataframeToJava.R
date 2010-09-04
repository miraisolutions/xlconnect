# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


dataframeToJava <- function(df) {
	# Force data.frame
	if(!is.data.frame(df)) {
		warning("Supplied object is not a data.frame! Trying to continue by converting it to a data.frame.")
		df <- as.data.frame(df)
	}
	
	dFrame <- jTryCatch(new(J("com.miraisolutions.xlconnect.integration.r.RDataFrameWrapper")))
	cnames <- colnames(df)
	for(i in seq(along = df)) {
		v <- df[, i]
		
		if(is(v, "numeric")) {
			jTryCatch(dFrame$addNumericColumn(cnames[i], .jarray(as.double(v)), .jarray(is.na(v))))
		}
		else if(is(v, "logical")) {
			jTryCatch(dFrame$addBooleanColumn(cnames[i], .jarray(v), .jarray(is.na(v))))
		}
		else if(is(v, "character")) {
			jTryCatch(dFrame$addStringColumn(cnames[i], .jarray(v), .jarray(is.na(v))))
		}
		else if(is(v, "factor")) {
			v <- as.character(v)
			jTryCatch(dFrame$addStringColumn(cnames[i], .jarray(v), .jarray(is.na(v))))
		}
		else if(is(v, "Date") || is(v, "POSIXt")) {
			v <- format(v, format = options("XLConnect.dateTimeFormat")[[1]])
			jTryCatch(dFrame$addDateTimeColumn(cnames[i], .jarray(v), .jarray(is.na(v))))
		}
		else
			stop("Unsupported data type (class) detected!")			
	}
	
	dFrame
}
