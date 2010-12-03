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
# Reading data from Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("readWorksheet",
	function(object, sheet, startRow, startCol, endRow, endCol, header) standardGeneric("readWorksheet"))


setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "logical"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			# returns a list of RDataFrameWrapper Java object references)
			dataFrame <- xlcCall(object@jobj$readWorksheet, as.integer(sheet - 1), as.integer(startRow - 1), 
				as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1), header, SIMPLIFY = FALSE)
			# construct data.frame
			dataFrame <- lapply(dataFrame, dataframeFromJava)
			
			# Return data.frame directly in case only one data.frame is read
			if(length(dataFrame) == 1) dataFrame[[1]]
			else dataFrame
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "missing"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			callGeneric(object, sheet, startRow, startCol, endRow, endCol, TRUE)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "logical"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			# returns a list of RDataFrameWrapper Java object references)
			dataFrame <- xlcCall(object@jobj$readWorksheet, sheet, as.integer(startRow - 1), as.integer(startCol - 1), 
				as.integer(endRow - 1), as.integer(endCol - 1), header, SIMPLIFY = FALSE)
			# construct data.frame
			dataFrame <- lapply(dataFrame, dataframeFromJava)
			
			# Return data.frame directly in case only one data.frame is read
			if(length(dataFrame) == 1) dataFrame[[1]]
			else dataFrame
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character", startRow = "numeric", startCol = "numeric", 
				endRow = "numeric", endCol = "numeric", header = "missing"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			callGeneric(object, sheet, startRow, startCol, endRow, endCol, TRUE)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "logical"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			# returns a list of RDataFrameWrapper Java object references)
			dataFrame <- xlcCall(object@jobj$readWorksheet, as.integer(sheet - 1), header, SIMPLIFY = FALSE)
			# construct data.frame
			dataFrame <- lapply(dataFrame, dataframeFromJava)
			
			# Return data.frame directly in case only one data.frame is read
			if(length(dataFrame) == 1) dataFrame[[1]]
			else dataFrame
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "missing"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			callGeneric(object, sheet, header = TRUE)
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "logical"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {	
			# returns a list of RDataFrameWrapper Java object references)
			dataFrame <- xlcCall(object@jobj$readWorksheet, sheet, header, SIMPLIFY = FALSE)
			# construct data.frame
			dataFrame <- lapply(dataFrame, dataframeFromJava)
			
			# Return data.frame directly in case only one data.frame is read
			if(length(dataFrame) == 1) dataFrame[[1]]
			else dataFrame
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character", startRow = "missing", startCol = "missing", 
				endRow = "missing", endCol = "missing", header = "missing"), 
		function(object, sheet, startRow, startCol, endRow, endCol, header) {
			callGeneric(object, sheet, header = TRUE)
		}
)
