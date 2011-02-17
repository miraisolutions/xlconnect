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
# Writing data to worksheets of an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("writeWorksheet",
	function(object, data, sheet, startRow, startCol, header) standardGeneric("writeWorksheet"))

# all args, sheet = num
setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			# pass data.frame's to Java - construct RDataFrameWrapper Java object references
			data <- lapply(wrapList(data), dataframeToJava)
			xlcCall(object@jobj$writeWorksheet, data, as.integer(sheet - 1), as.integer(startRow - 1),
				as.integer(startCol - 1), header)
			invisible()
		}
)

# all args, sheet = char
setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			# pass data.frame's to Java - construct RDataFrameWrapper Java object references
			data <- lapply(wrapList(data), dataframeToJava)
			xlcCall(object@jobj$writeWorksheet, data, sheet, as.integer(startRow - 1), 
				as.integer(startCol - 1), header)
			invisible()
		}
)
# no header, sheet = char
setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "numeric", 
				startCol = "numeric", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			callGeneric(object, data, sheet, startRow, startCol, TRUE)
		}
)
# no header, sheet = num
setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			# pass data.frame's to Java - construct RDataFrameWrapper Java object references
			data <- lapply(wrapList(data), dataframeToJava)
			xlcCall(object@jobj$writeWorksheet, data, as.integer(sheet - 1), header)
			invisible()
		}
)

# no coords, sheet=char, w/ header
setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			# pass data.frame's to Java - construct RDataFrameWrapper Java object references
			data <- lapply(wrapList(data), dataframeToJava)
			xlcCall(object@jobj$writeWorksheet, data, sheet, header)
			invisible()
		}
)

# no coords, sheet=num, w/ header
setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "logical"), 
		function(object, data, sheet, startRow, startCol, header) {
			# pass data.frame's to Java - construct RDataFrameWrapper Java object references
			data <- lapply(wrapList(data), dataframeToJava)
			xlcCall(object@jobj$writeWorksheet, data, as.integer(sheet - 1), header)
			invisible()
		}
)
          
# no coords, sheet=num, w/o header 
setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "numeric", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			callGeneric(object, data, sheet, header = TRUE)
		}
)

# no coords, sheet=char, w/o header 
setMethod("writeWorksheet", 
		signature(object = "workbook", data = "ANY", sheet = "character", startRow = "missing", 
				startCol = "missing", header = "missing"), 
		function(object, data, sheet, startRow, startCol, header) {
			callGeneric(object, data, sheet, header = TRUE)
		}
)
