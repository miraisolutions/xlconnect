#############################################################################
#
# XLConnect
# Copyright (C) 2010 Mirai Solutions GmbH
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
# TODO: Add Comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("writeWorksheet",
	function(object, data, sheet, startRow, startCol, header) standardGeneric("writeWorksheet"))

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			# note that Java indices are 0-based
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), as.integer(sheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), sheet, as.integer(startRow - 1), 
					as.integer(startCol - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			callGeneric(object, data, sheet, startRow, startCol, TRUE)
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), as.integer(sheet - 1), header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			jTryCatch(object@jobj$writeWorksheet(dataframeToJava(data), sheet, header))
			invisible()
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			callGeneric(object, data, sheet, header = TRUE)
		}
)

setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			callGeneric(object, data, sheet, header = TRUE)
		}
)
