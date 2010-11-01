# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setMethod("$", 
		signature(x = "cellstyle"),
		function(x, name) {
			g <- getGeneric(name)
			if(is.null(g)) stop("Method undefined for class 'cellstyle'")
			function(...) g(x, ...)
		}
)
