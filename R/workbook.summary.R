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
# Outputs a summary of workbook object
#
# The summary includes:
#	- the underlying filename
#	- contained worksheets
#	- hidden sheets
#	- very hidden sheets
#	- defined names
#	- active sheet name
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setMethod("summary", signature(object = "workbook"), function(object) {
	cat("*** XLConnect Workbook Summary ***\n")
	cat("> Filename: '", object@filename, "'\n", sep = "")
	
	nice <- function(x) {
		if(length(x) > 0 && !is.na(x)) x
		else "<NONE>"
	}
	
	sheets <- getSheets(object)
	
	cat("> Sheets (all):\n")
	cat(nice(sheets), sep = ", ", fill = TRUE)
	
	cat("> Hidden Sheets:\n")
	idx <- as.logical(sapply(sheets, function(s) isSheetHidden(object, s)))
	cat(nice(sheets[idx]), sep = ", ", fill = TRUE)
	
	cat("> Very Hidden Sheets:\n")
	idx <- as.logical(sapply(sheets, function(s) isSheetVeryHidden(object, s)))
	cat(nice(sheets[idx]), sep = ", ", fill = TRUE)
	
	cat("> Names:\n")
	cat(nice(getDefinedNames(object)), sep = ", ", fill = TRUE)
	
	cat("> Active Sheet: ", nice(getActiveSheetName(object)), "\n")
})
