# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


normalizeDataframe <- function(df) {
	rownames(df) <- NULL
	res <- lapply(df, 
			function(col) {
				if(is(col, "factor")) {
					as.character(col)
				} else if(is(col, "Date") || is(col, "POSIXt")) {
					# Get rid of "original" timezone and assume UTC
					as.POSIXct(
						format(col, format = options("XLConnect.dateTimeFormat")[[1]]), 
						tz = "UTC")
				} else
					col
			})
	
	data.frame(res, stringsAsFactors = F)
}
