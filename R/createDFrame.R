# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


createDFrame <- function(df) {
	# TODO: better error message required ...
	if(!is.data.frame(df)) stop("Only data.frame's are allowed!")
	dFrame <- new("dframe")
	cnames <- colnames(df)
	for(i in seq(along = df)) {
		res <- prepareForXLConnect(df[, i])
		addColumn(dFrame, name = cnames[i], dataType = res$dataType, data = res$data)
	}
	
	dFrame
}
