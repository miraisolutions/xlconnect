# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("removeName")) {
	if(is.function("removeName")) fun <- removeName
	else fun <- function(.Object, name) standardGeneric("removeName")
	setGeneric("removeName", fun)
}

setMethod("removeName", 
		signature(.Object = "workbook", name = "character"), 
		function(.Object, name) {
			.Object@jobj$removeName(name)
		}
)
