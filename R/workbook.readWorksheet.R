#############################################################################
#
# XLConnect
# Copyright (C) 2010-2012 Mirai Solutions GmbH
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
	function(object, sheet, ...) standardGeneric("readWorksheet"))

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, startRow = 0, startCol = 0, endRow = 0, endCol = 0, region = NULL,
				 header = TRUE, rownames = NULL, colTypes = character(0), forceConversion = FALSE, 
				 dateTimeFormat = getOption("XLConnect.dateTimeFormat")) {
			 
			if(!is.null(region)) {
				# Convert region to indices
				idx = rg2idx(region)
				startRow = idx[,1]
				startCol = idx[,2]
				endRow = idx[,3]
				endCol = idx[,4]
			}
			
			# returns a list of RDataFrameWrapper Java object references)
			dataFrame <- xlcCall(object, "readWorksheet", as.integer(sheet - 1), as.integer(startRow - 1), 
				as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1), header, 
				.jarray(classToXlcType(colTypes)), forceConversion, dateTimeFormat, SIMPLIFY = FALSE)
			# construct data.frame
			dataFrame <- lapply(dataFrame, function(x) {
				extractRownames(dataframeFromJava(x), rownames)
			})
			
			# Return data.frame directly in case only one data.frame is read
			if(length(dataFrame) == 1) dataFrame[[1]]
			else dataFrame
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet, startRow = 0, startCol = 0, endRow = 0, endCol = 0, region = NULL,
				 header = TRUE, rownames = NULL, colTypes = character(0), forceConversion = FALSE, 
				 dateTimeFormat = getOption("XLConnect.dateTimeFormat")) {
			
			sheet = getSheetPos(object, sheet)
			callGeneric()
		}
)
