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
				} else
					col
			})
	
	data.frame(res, stringsAsFactors = F)
}
