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
# Utility to get a subset of columns (with numeric indices) while using
# keep or drop parameters in readNamedRegion or readWorksheet
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

getColSubset <- function(object, sheet, startRow, endRow, startCol, endCol, header, numcols, keep, drop){
	if (!is.null(keep) && !is.null(drop)) {
		stop("Specify either keep OR drop (not both)")
	} else { 
		if (is.null(drop)){
			if(!is.list(keep)) keep = list(keep)
			charColKept <- unlist(lapply(keep, is.character))
			if(any(charColKept) && !header)
				stop(paste("Header=F and keep values", paste(unlist(keep)[charColKept], collapse=", "), "of type character", sep=" "))
			subset = lapply(keep, function(kp) {
						if (is.numeric(kp)) {
							outerCols=setdiff(kp, seq(numcols))
							if (length(outerCols)!=0) {
								stop(sprintf("Column(s) '%s' out of the bounding box!", paste(outerCols, collapse = ", ")))
							} else {
								subset = kp
							}
						} else
						if (is.character(kp)) {
							headerdf <- readWorksheet(object, sheet, startRow = startRow, endRow = startRow, startCol = startCol, endCol = endCol, header = FALSE)
							columnNames = unlist(headerdf)
							idx = match(kp, columnNames)
							subset = idx
							idx.na = is.na(idx)
							if(any(idx.na)) {
								stop(sprintf("Column name(s) '%s' not existing or out of the bounding box!", paste(kp[idx.na], collapse = ", ")))
							}
						}
						
						.jarray(as.integer(subset-1))
					})
		} else {
			if(!is.list(drop)) drop = list(drop)
			charColDropped <- unlist(lapply(drop, is.character))
			if(any(charColDropped) && !header)
				stop(paste("Header=F and drop values", paste(unlist(drop)[charColDropped], collapse=", "), "of type character", sep=" "))
			drop = rep(drop, len=length(numcols))
			subset = lapply(1:length(numcols), function(dpel) {
						if (is.numeric(drop[[dpel]])) {
							outerCols=setdiff(drop[[dpel]], seq(numcols[dpel]))
							if (length(outerCols)!=0) {
								stop(sprintf("Column(s) '%s' out of the bounding box!", paste(outerCols, collapse = ", ")))
							} else {
								subset = setdiff(seq(numcols[dpel]), drop[[dpel]])
							}
						} else if (is.character(drop[[dpel]])) {
							if(header==F)
								stop(paste("Error: Header=F and drop value", drop[[dpel]], "of type character"))
							headerdf <- readWorksheet(object, sheet, startRow = startRow, endRow = startRow, startCol = startCol, endCol = endCol, header = FALSE)
							columnNames = unlist(headerdf)
							idx = match(drop[[dpel]], columnNames)
							subset = setdiff(seq(numcols[[dpel]]), idx)
							idx.na = is.na(idx)
							if(any(idx.na)) {
								stop(sprintf("Column name(s) '%s' not existing or out of the bounding box!", paste(drop[[dpel]][idx.na], collapse = ", ")))
							}
						}
						.jarray(as.integer(subset-1))
					})
		}
	}	
	subset
}