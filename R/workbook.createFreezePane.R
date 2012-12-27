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
# Creating freeze panes
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

setGeneric("createFreezePane",
	function(object, sheet, colSplit, rowSplit, leftColumn = colSplit, topRow = rowSplit) standardGeneric("createFreezePane"))

setMethod("createFreezePane", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, colSplit, rowSplit, leftColumn, topRow) {
			if (is.numeric(colSplit)) {
				colSplit = colSplit
			} else if (is.character(colSplit)){
				colSplit = col2idx(colSplit)
			} else {
				stop(sprintf("colSplit should be numeric or character"))
			}
			
			if (is.numeric(leftColumn)) {
				leftColumn = leftColumn
			} else if (is.character(leftColumn)){
				leftColumn = col2idx(leftColumn)
			} else {
				stop(sprintf("leftColumn should be numeric or character"))
			}
			xlcCall(object, "createFreezePane", as.integer(sheet - 1), as.integer(colSplit-1), as.integer(rowSplit-1), as.integer(leftColumn-1), as.integer(topRow-1))
			invisible()	
		}
)

setMethod("createFreezePane", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet, colSplit, rowSplit, leftColumn, topRow) {
			if (is.numeric(colSplit)) {
				colSplit = colSplit
			} else if (is.character(colSplit)){
				colSplit = col2idx(colSplit)
			} else {
				stop(sprintf("colSplit should be numeric or character"))
			}
			
			if (is.numeric(leftColumn)) {
				leftColumn = leftColumn
			} else if (is.character(leftColumn)){
				leftColumn = col2idx(leftColumn)
			} else {
				stop(sprintf("leftColumn should be numeric or character"))
			}
			xlcCall(object, "createFreezePane", sheet, as.integer(colSplit-1), as.integer(rowSplit-1), as.integer(leftColumn-1), as.integer(topRow-1))
			invisible()
		}
)