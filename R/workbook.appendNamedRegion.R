#############################################################################
#
# XLConnect
# Copyright (C) 2010-2025 Mirai Solutions GmbH
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Appending data to named regions
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("appendNamedRegion",
	function(object, data, name, header = FALSE, overwriteFormulaCells = TRUE,
	         rownames = NULL, worksheetScope = NULL) standardGeneric("appendNamedRegion"))

setMethod("appendNamedRegion", 
	signature(object = "workbook", data = "ANY"), 
	function(object, data, name, header = FALSE, overwriteFormulaCells = TRUE,
	         rownames = NULL, worksheetScope = NULL) {
		if(is.character(rownames))
			data <- includeRownames(data, rownames)
		# pass data.frame's to Java - construct RDataFrameWrapper Java object references
		data <- lapply(wrapList(data), dataframeToJava)
		xlcCall(object, "appendNamedRegion", data, name, header, overwriteFormulaCells, worksheetScope %||% .jnull(), .simplify = FALSE)
		invisible()
	}
)
