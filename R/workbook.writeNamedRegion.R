# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

if(!isGeneric("writeNamedRegion")) {
	if(is.function("writeNamedRegion")) fun <- getSheets
	else fun <- function(.Object, data, name, location, overwrite) standardGeneric("writeNamedRegion")
	setGeneric("writeNamedRegion", fun)
}

setMethod("writeNamedRegion", 
	signature(.Object = "workbook", data = "ANY", name = "character", location = "character", overwrite = "logical"), 
	function(.Object, data, name, location, overwrite) {
	
	if(!is.data.frame(data)) data <- as.data.frame(data)
	.Object@jobj$writeNamedRegion(createDFrame(data)@jobj, name, location, overwrite)
	
	invisible()
})

setMethod("writeNamedRegion", 
	signature(.Object = "workbook", data = "ANY", name = "character", location = "missing", overwrite = "missing"), 
	function(.Object, data, name, location, overwrite) {
	
	if(!is.data.frame(data)) data <- as.data.frame(data)
	.Object@jobj$writeNamedRegion(createDFrame(data)@jobj, name)
	
	invisible()
})
