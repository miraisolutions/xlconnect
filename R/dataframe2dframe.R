# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


dataframe2dframe <- function(df) {
	# TODO: better error message required ...
	if(!is.data.frame(df)) stop("Only data.frame's are allowed!")
	
	dFrame <- new(J("com.miraisolutions.xlconnect.integration.r.RDataFrameWrapper"))
	cnames <- colnames(df)
	for(i in seq(along = df)) {
		res <- prepareForXLConnect(df[, i])
		
		data <- res$data
		switch(res$dataType,
				
				"Numeric" = {
					dFrame$addNumericColumn(cnames[i], .jarray(data), .jarray(is.na(data)))
				},
				
				"Boolean" = {
					dFrame$addBooleanColumn(cnames[i], .jarray(data), .jarray(is.na(data)))
				},
				
				"String" = {
					dFrame$addStringColumn(cnames[i], .jarray(data), .jarray(is.na(data)))
				},
				
				"DateTime" = {
					dFrame$addDateTimeColumn(cnames[i], .jarray(data), .jarray(is.na(data)))
				},
				
				stop("Unsupported data type!")
		)			
	}
	
	dFrame
}
