# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setGeneric("createCellStyle",
		function(object, name) standardGeneric("createCellStyle"))

setMethod("createCellStyle", 
		signature(object = "workbook", name = "character"), 
		function(object, name) {
			jobj <- jTryCatch(object@jobj$createCellStyle(name))
			new("cellstyle", jobj = jobj)
		}
)

setMethod("createCellStyle", 
		signature(object = "workbook", name = "missing"), 
		function(object, name) {
			jobj <- jTryCatch(object@jobj$createCellStyle())
			new("cellstyle", jobj = jobj)
		}
)
