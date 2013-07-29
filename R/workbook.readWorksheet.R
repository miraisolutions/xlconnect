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
# Reading data from Excel worksheets
# 
# Authors:  Martin Studer, Mirai Solutions GmbH
#           Thomas Themel, Mirai Solutions GmbH
#           Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

setGeneric("readWorksheet",
	function(object, sheet, startRow = 0, startCol = 0, endRow = 0, endCol = 0,
           autofitRow = TRUE, autofitCol = TRUE, region = NULL,
           header = TRUE, rownames = NULL, colTypes = character(0), forceConversion = FALSE, 
           dateTimeFormat = getOption("XLConnect.dateTimeFormat"), check.names = TRUE, 
           useCachedValues = FALSE, keep = NULL, drop = NULL, simplify = FALSE,
           readStrategy = "default") 
    
    standardGeneric("readWorksheet"))

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, startRow = 0, startCol = 0, endRow = 0, endCol = 0, 
             autofitRow = TRUE, autofitCol = TRUE, region = NULL, header = TRUE, 
             rownames = NULL, colTypes = character(0), forceConversion = FALSE, 
             dateTimeFormat = getOption("XLConnect.dateTimeFormat"), check.names = TRUE, 
             useCachedValues = FALSE, keep = NULL, drop = NULL, simplify = FALSE,
             readStrategy = "default") {
			 
			if(!is.null(region)) {
				# Convert region to indices
				idx = rg2idx(region)
				startRow = idx[,1]
				startCol = idx[,2]
				endRow = idx[,3]
				endCol = idx[,4]
			}

			boundingBoxDim = getBoundingBox(object, sheet, startRow = startRow, startCol = startCol, endRow = endRow, endCol = endCol,
		                                  autofitRow = autofitRow, autofitCol = autofitCol)
			startRow = boundingBoxDim[1, ]
		  startCol = boundingBoxDim[2, ]
		 	endRow = boundingBoxDim[3, ]
		  endCol = boundingBoxDim[4, ]
		  numcols = ifelse(endCol == 0, 0, endCol - startCol + 1)
			
			subset <- getColSubset(object, sheet, startRow, endRow, startCol, endCol, header, numcols, keep, drop)
			# returns a list of RDataFrameWrapper Java object references)
			dataFrame <- xlcCall(object, "readWorksheet", as.integer(sheet - 1), as.integer(startRow - 1), 
					as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1), header, 
					.jarray(classToXlcType(colTypes)), forceConversion, dateTimeFormat, useCachedValues, subset, 
          readStrategy, SIMPLIFY = FALSE)
	
			# get data.frames from Java
			dataFrame = lapply(dataFrame, dataframeFromJava, check.names = check.names)
			# extract rownames
			dataFrame = extractRownames(dataFrame, rownames)
      
      # simplify
      dataFrame =
			mapply(df = dataFrame, simplify = rep(simplify, length.out = length(dataFrame)), 
			       FUN = function(df, simplify) {
			         if(simplify) unlist(df, use.names = FALSE)
			         else df
			       }, SIMPLIFY = FALSE
			)
			
			# Return data.frame directly in case only one data.frame is read
			if(length(dataFrame) == 1) dataFrame[[1]]
			else dataFrame
		}
)

setMethod("readWorksheet", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet, startRow = 0, startCol = 0, endRow = 0, endCol = 0,
             autofitRow = TRUE, autofitCol = TRUE, region = NULL, header = TRUE, 
             rownames = NULL, colTypes = character(0), forceConversion = FALSE, 
             dateTimeFormat = getOption("XLConnect.dateTimeFormat"), check.names = TRUE, 
             useCachedValues = TRUE, keep = NULL, drop = NULL, simplify = FALSE,
             readStrategy = "default") {
			
			# Remember sheet names
			sheetNames = sheet
			sheet = getSheetPos(object, sheet)
			# Get result from calling generic with sheet positions
			dataFrame = callGeneric()
			# If more than one worksheet is read the results will be a list
			# --> assign names of sheets to list
			if(length(sheet) > 1)
				names(dataFrame) = sheetNames
			
			dataFrame
		}
)
