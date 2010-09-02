# Wrapper to tryCatch for standard Java exception handling
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

jTryCatch <- function(...) {
	tryCatch(..., Throwable = 
		function(e) {
			if(!is.jnull(e$jobj)) {
				stop(paste(class(e)[1], e$jobj$getMessage(), sep = " (Java): "),
					call. = FALSE)
			} else 
				stop("Undefined error occurred!")
		}
	)
}
