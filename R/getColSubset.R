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
# Utility to get a subset of columns (with numeric indices) while using
# keep or drop parameters in readNamedRegion or readWorksheet
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

getColSubset <- function(object, sheet, startRow, endRow, startCol, endCol, header, numcols, keep, drop){
	globalSubset <- list()
	args <-	list(sheet=sheet, 
			startRow=startRow, 
			endRow=endRow, 
			startCol=startCol, 
			endCol=endCol, 
			header=header, 
			numcols=numcols, 
			keep=if(is(keep, "list")) keep else list(keep), 
			drop=if(is(drop, "list")) drop else list(drop))
	maxlen <- max(sapply(args, length))
	args_repl <- lapply(args, function(x) rep(x, length = maxlen))
	
	for(i in seq_len(maxlen)) {
		cargs = lapply(args_repl, "[[", i)
		if(is.null(cargs$drop) && is.null(cargs$keep)){
			subset = seq_len(cargs$numcols)
			globalSubset[[i]] <- .jarray(as.integer(subset-1))
		} else if(!is.null(cargs$drop) && !is.null(cargs$keep)) {
			stop("Specify either keep OR drop (not both)")
		} else {
			if (is.null(cargs$drop)){
				if(is.character(cargs$keep) && !cargs$header)
					stop(paste("Header=F and keep values", paste(cargs$keep, collapse=", "), "of type character", sep=" "))	
				
				if (is.numeric(cargs$keep)) {
					outerCols=setdiff(cargs$keep, seq_len(cargs$numcols))
					if (length(outerCols)!=0) {
						stop(sprintf("Column(s) '%s' out of the bounding box!", paste(outerCols, collapse = ", ")))
					} else {
						subset = cargs$keep
					}
				} else
				if (is.character(cargs$keep)) {
					columnNames <- readWorksheet(object, cargs$sheet, startRow = cargs$startRow, endRow = cargs$startRow, startCol = cargs$startCol, endCol = cargs$endCol, header = FALSE)
					idx = match(cargs$keep, columnNames)
					subset = idx
					idx.na = is.na(idx)
					if(any(idx.na)) {
						stop(sprintf("Column name(s) '%s' not existing or out of the bounding box!", paste(cargs$keep[idx.na], collapse = ", ")))
					}
				}
				
				globalSubset[[i]] <- .jarray(as.integer(subset-1))
		
			} else {
				if(is.character(cargs$drop) && !cargs$header)
					stop(paste("Header=F and drop values", paste(cargs$drop, collapse=", "), "of type character", sep=" "))
				if (is.numeric(cargs$drop)) {
					outerCols=setdiff(cargs$drop, seq_len(cargs$numcols))
					if (length(outerCols)!=0) {
						stop(sprintf("Column(s) '%s' out of the bounding box!", paste(outerCols, collapse = ", ")))
					} else {
						subset = setdiff(seq_len(cargs$numcols), cargs$drop)
					}
				} else if (is.character(cargs$drop)) {
					columnNames <- readWorksheet(object, cargs$sheet, startRow = cargs$startRow, endRow = cargs$startRow, startCol = cargs$startCol, endCol = cargs$endCol, header = FALSE)
					idx = match(cargs$drop, columnNames)
					subset = setdiff(seq_len(cargs$numcols), idx)
					idx.na = is.na(idx)
					if(any(idx.na)) {
						stop(sprintf("Column name(s) '%s' not existing or out of the bounding box!", paste(cargs$drop[idx.na], collapse = ", ")))
					}
				}
				globalSubset[[i]] <- .jarray(as.integer(subset-1))

			}
		}
	}
	
	globalSubset
}