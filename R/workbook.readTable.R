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
# Reading tables from an Excel file
# 
# Authors:  Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("readTable",
	function(object, sheet, table, header = TRUE, rownames = NULL, colTypes = character(0),
			forceConversion = FALSE, dateTimeFormat = getOption("XLConnect.dateTimeFormat"),
			check.names = TRUE, useCachedValues = FALSE, keep = NULL, drop = NULL, simplify = FALSE,
      readStrategy = "default") 
    standardGeneric("readTable"))


setMethod("readTable", 
	signature(object = "workbook", sheet = "numeric"), 
	function(object, sheet, table, header = TRUE, rownames = NULL, colTypes = character(0),
			forceConversion = FALSE, dateTimeFormat = getOption("XLConnect.dateTimeFormat"),
			check.names = TRUE, useCachedValues = FALSE, keep = NULL, drop = NULL, simplify = FALSE,
      readStrategy = "default") {

		# returns a list of RDataFrameWrapper Java object references
		namedim = matrix(as.vector(t(getReferenceCoordinatesForTable(object, sheet, table))), nrow = 4, byrow = FALSE)
		startRow = namedim[1,]
		startCol = namedim[2,]
		endRow = namedim[3,]
		endCol = namedim[4,]
		numcols = endCol - startCol + 1

		subset <- getColSubset(object, sheet, startRow, endRow, startCol, endCol, header, numcols, keep, drop)
		dataFrame <- xlcCall(object, "readTable", as.integer(sheet - 1), table, header, .jarray(classToXlcType(colTypes)), 
				forceConversion, dateTimeFormat, useCachedValues, subset, readStrategy, SIMPLIFY = FALSE)
		

    # get data.frames from Java
    dataFrame = lapply(dataFrame, dataframeFromJava, check.names = check.names)
		# extract rownames
    dataFrame = extractRownames(dataFrame, rownames)
		names(dataFrame) <- table
    
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

setMethod("readTable", 
  signature(object = "workbook", sheet = "character"), 
  function(object, sheet, table, header = TRUE, rownames = NULL, colTypes = character(0),
           forceConversion = FALSE, dateTimeFormat = getOption("XLConnect.dateTimeFormat"),
           check.names = TRUE, useCachedValues = FALSE, keep = NULL, drop = NULL, simplify = FALSE,
           readStrategy = "default") {
    
    # Remember sheet names
    sheet = getSheetPos(object, sheet)
    # Get result from calling generic with sheet positions
    dataFrame = callGeneric()
    # If more than one worksheet is read the results will be a list
    # --> assign names of sheets to list
    if(length(table) > 1)
      names(dataFrame) = table
    
    dataFrame    
  }
)