# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

dataframeFromJava <- function(df) {
	
	if(!is(df, "jobjRef"))
		stop("Invalid object - object of class 'jobjRef' required!")
	
	columnTypes <- jTryCatch(df$getColumnTypes())
	columnNames <- jTryCatch(df$getColumnNames())
	
	# Init result list to contain column vectors
	res <- list()
	for(i in seq(along = columnTypes)) {
		# Note, Java indices are 0-based while R's are 1-based ...
		jIndex <- as.integer(i - 1)
		
		switch(columnTypes[i],
				
				"Numeric" = {
					res[[i]] <- as.vector(jTryCatch(df$getNumericColumn(jIndex)))
				},
				
				"String" = {
					res[[i]] <- as.vector(jTryCatch(df$getStringColumn(jIndex)))
				},
				
				"Boolean" = {
					res[[i]] <- as.vector(jTryCatch(df$getBooleanColumn(jIndex)))
				},
				
				"DateTime" = {
					# Convert date/time strings back to POSIXct with timezone UTC
					res[[i]] <- as.POSIXct(as.vector(jTryCatch(df$getDateTimeColumn(jIndex))), 
							format = options("XLConnect.dateTimeFormat")[[1]], tz = "UTC")
				},
				
				stop("Unsupported column type detected!")
		)
		
		# Put missings back in place;
		# note that Java primitives are communicated back which 
		# don't support any missings - therefore this step
		res[[i]][jTryCatch(df$isMissing(jIndex))] <- NA
	}
	
	# Apply names
	names(res) <- columnNames
	
	data.frame(res, stringsAsFactors = FALSE)
}
