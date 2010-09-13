# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setMethod("$", 
	signature(x = "workbook"),
	function(x, name) {
		g <- getGeneric(name)
		if(is.null(g)) stop("Method undefined for class 'workbook'")
		function(...) g(x, ...)
	}
)
