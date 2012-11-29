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
# Reading named regions from an Excel file
# 
# Authors:  Martin Studer, Mirai Solutions GmbH
#           Thomas Themel, Mirai Solutions GmbH
#           Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

setGeneric("readNamedRegion",
	function(object, name, header = TRUE, rownames = NULL, colTypes = character(0),
			forceConversion = FALSE, dateTimeFormat = getOption("XLConnect.dateTimeFormat"),
			check.names = TRUE, useCachedValues = FALSE, keep=NULL, drop=NULL) standardGeneric("readNamedRegion"))


setMethod("readNamedRegion", 
	signature(object = "workbook"), 
	function(object, name, header = TRUE, rownames = NULL, colTypes = character(0),
			forceConversion = FALSE, dateTimeFormat = getOption("XLConnect.dateTimeFormat"),
			check.names = TRUE, useCachedValues = FALSE, keep=NULL, drop=NULL) {

		# returns a list of RDataFrameWrapper Java object references
			
		sheet = as.vector(extractSheetName(getReferenceFormula(object, name)))
		namedim = getReferenceCoordinates(object, name)
		startRow = namedim[1,1]
		startCol = namedim[1,2]
		endCol = namedim[2,2]
		numcols = endCol-startCol+1

		
		if (is.null(keep) && is.null(drop)) {
			dataFrame <- xlcCall(object, "readNamedRegion", name, header, .jarray(classToXlcType(colTypes)), 
					forceConversion, dateTimeFormat, useCachedValues, .jnull("java/lang/Integer"), SIMPLIFY = FALSE)
		} else
		if (!is.null(keep) && !is.null(drop)) {
			stop(sprintf("Specify either keep OR drop (not both)"))
		} else { 
			if (is.null(drop)){
				if (is.numeric(keep)) {
					outerCols=setdiff(keep, seq(numcols))
					if (length(outerCols)!=0) {
						stop(sprintf("Column(s) '%s' out of the named region!", paste(outerCols, collapse = ", ")))
					} else {
						subset = keep
					}
				} else
				if (is.character(keep)) {
					headerdf <- readWorksheet(object, sheet, startRow = startRow, endRow = startRow, startCol = startCol, endCol = endCol, header = FALSE)
					columnNames = unlist(headerdf)
					idx = match(keep, columnNames)
					idx.na = is.na(idx)
					if(any(idx.na)) {
						stop(sprintf("Column name(s) '%s' not existing or out of the named region!", paste(keep[idx.na], collapse = ", ")))
					}
					subset = idx
				}
			} else {
				if (is.numeric(drop)) {
					outerCols=setdiff(drop, seq(numcols))
					if (length(outerCols)!=0) {
						stop(sprintf("Column(s) '%s' out of the named region!", paste(outerCols, collapse = ", ")))
					} else {
						subset = setdiff(seq(numcols), drop)
					}
				} else
				if (is.character(drop)) {
					headerdf <- readWorksheet(object, sheet, startRow = startRow, endRow = startRow, startCol = startCol, endCol = endCol, header = FALSE)
					columnNames = unlist(headerdf)
					idx = match(drop, columnNames)
					subset = setdiff(seq(numcols), idx)
					idx.na = is.na(idx)
					if(any(idx.na)) {
						stop(sprintf("Column name(s) '%s' not existing or out of the named region!", paste(drop[idx.na], collapse = ", ")))
					}
				}
			}
			dataFrame <- xlcCall(object, "readNamedRegion", name, header, .jarray(classToXlcType(colTypes)), 
					forceConversion, dateTimeFormat, useCachedValues, .jarray(as.integer(subset-1)), SIMPLIFY = FALSE)
		}

    # get data.frames from Java
    dataFrame = lapply(dataFrame, dataframeFromJava, check.names = check.names)
		# extract rownames
    dataFrame = extractRownames(dataFrame, rownames)
		names(dataFrame) <- name
		
		# Return data.frame directly in case only one data.frame is read
		if(length(dataFrame) == 1) dataFrame[[1]]
		else dataFrame
	}
)
