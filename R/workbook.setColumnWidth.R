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
# Setting the width of a column in a worksheet
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setColumnWidth",
		function(object, sheet, column, width) standardGeneric("setColumnWidth"))

setMethod("setColumnWidth", 
		signature(object = "workbook", sheet = "numeric", column = "numeric", width = "numeric"), 
		function(object, sheet, column, width) {
			xlcCall(object@jobj$setColumnWidth, as.integer(sheet - 1), as.integer(column - 1),
				as.integer(width))
			invisible()
		}
)

setMethod("setColumnWidth", 
		signature(object = "workbook", sheet = "character", column = "numeric", width = "numeric"), 
		function(object, sheet, column, width) {
			xlcCall(object@jobj$setColumnWidth, sheet, as.integer(column - 1),
				as.integer(width))
			invisible()
		}
)