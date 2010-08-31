# Wrapper to tryCatch for standard Java exception handling
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

jTryCatch <- function(...) {
	tryCatch(..., Exception = 
		function(e) {
			stop(paste(class(e)[1], e$jobj$getMessage(), sep = ": "),
				call. = FALSE)
		}
	)
}
