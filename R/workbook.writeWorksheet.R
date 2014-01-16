#############################################################################
#
# XLConnect
# Copyright (C) 2010-2013 Mirai Solutions GmbH
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
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Writing data to worksheets of an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("writeWorksheet",
	function(object, data, sheet, startRow = 1, startCol = 1, header = TRUE, rownames = NULL) standardGeneric("writeWorksheet"))

setMethod("writeWorksheet", 
	signature(object = "workbook", data = "ANY", sheet = "numeric"), 
	function(object, data, sheet, startRow = 1, startCol = 1, header = TRUE, rownames = NULL) {
		data <- includeRownames(data, rownames)
		# pass data.frame's to Java - construct RDataFrameWrapper Java object references
		data <- lapply(wrapList(data), dataframeToJava)
		xlcCall(object, "writeWorksheet", data, as.integer(sheet - 1), as.integer(startRow - 1),
			as.integer(startCol - 1), header)
		invisible()
	}
)

setMethod("writeWorksheet", 
	signature(object = "workbook", data = "ANY", sheet = "character"), 
	function(object, data, sheet, startRow = 1, startCol = 1, header = TRUE, rownames = NULL) {
	  data <- includeRownames(data, rownames)
		# pass data.frame's to Java - construct RDataFrameWrapper Java object references
		data <- lapply(wrapList(data), dataframeToJava)
		xlcCall(object, "writeWorksheet", data, sheet, as.integer(startRow - 1), 
			as.integer(startCol - 1), header)
		invisible()
	}
)
