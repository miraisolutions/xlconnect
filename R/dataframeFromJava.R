# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

dataframeFromJava <- function(df) {
	columnTypes <- df$getColumnTypes()
	columnNames <- df$getColumnNames()
	
	# Init result list to contain column vectors
	res <- list()
	for(i in seq(along = columnTypes)) {
		# Note, Java indices are 0-based while R's are 1-based ...
		jIndex <- as.integer(i - 1)
		
		switch(columnTypes[i],
				
				"Numeric" = {
					res[[i]] <- as.vector(df$getNumericColumn(jIndex))
				},
				
				"String" = {
					res[[i]] <- as.vector(df$getStringColumn(jIndex))
				},
				
				"Boolean" = {
					res[[i]] <- as.vector(df$getBooleanColumn(jIndex))
				},
				
				"DateTime" = {
					# Convert date/time strings back to POSIXlt
					res[[i]] <- as.POSIXlt(as.vector(df$getDateTimeColumn(jIndex)), format = options("XLConnect.dateTimeFormat")[[1]])
				},
				
				stop("Unsupported column type detected!")
		)
		
		# Put missings back in place;
		# note that Java primitives are communicated back which 
		# don't support any missings - therefore this step
		res[[i]][df$isMissing(jIndex)] <- NA
	}
	
	# Apply names
	names(res) <- columnNames
	
	data.frame(res, stringsAsFactors = F)
}

