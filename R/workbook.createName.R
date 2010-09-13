# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("createName",
	function(object, name, formula, overwrite) standardGeneric("createName"))

setMethod("createName", 
		signature(object = "workbook", name = "character", formula = "character", 
				overwrite = "logical"), 
		function(object, name, formula, overwrite) {
			jTryCatch(object@jobj$createName(name, formula, overwrite))
		}
)

setMethod("createName", 
		signature(object = "workbook", name = "character", formula = "character", 
				overwrite = "missing"), 
		function(object, name, formula, overwrite) {
			jTryCatch(object@jobj$createName(name, formula, FALSE))
		}
)
