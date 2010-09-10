# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("createName")) {
	if(is.function("createName")) fun <- createName
	else fun <- function(.Object, name, formula, overwrite) standardGeneric("createName")
	setGeneric("createName", fun)
}

setMethod("createName", 
		signature(.Object = "workbook", name = "character", formula = "character", 
				overwrite = "logical"), 
		function(.Object, name, formula, overwrite) {
			jTryCatch(.Object@jobj$createName(name, formula, overwrite))
		}
)

setMethod("createName", 
		signature(.Object = "workbook", name = "character", formula = "character", 
				overwrite = "missing"), 
		function(.Object, name, formula, overwrite) {
			jTryCatch(.Object@jobj$createName(name, formula, FALSE))
		}
)
