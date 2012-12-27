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
# Setting the height of a row in a worksheet
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setRowHeight",
		function(object, sheet, row, height = -1) standardGeneric("setRowHeight"))

setMethod("setRowHeight", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, row, height = -1) {
			xlcCall(object, "setRowHeight", as.integer(sheet - 1), as.integer(row - 1), height)
			invisible()
		}
)

setMethod("setRowHeight", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet, row, height = -1) {
			xlcCall(object, "setRowHeight", sheet, as.integer(row - 1), height)
			invisible()
		}
)
