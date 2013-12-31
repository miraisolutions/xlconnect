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
# Querying the coordinates of a worksheet bounding box
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

setGeneric("getBoundingBox",
		function(object, sheet, startRow = 0, startCol = 0, endRow = 0, endCol = 0,
             autofitRow = TRUE, autofitCol = TRUE) standardGeneric("getBoundingBox"))

setMethod("getBoundingBox", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, startRow = 0, startCol = 0, endRow = 0, endCol = 0,
             autofitRow = TRUE, autofitCol = TRUE) {
			res <- xlcCall(object, "getBoundingBox", as.integer(sheet - 1), as.integer(startRow - 1), 
                     as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1),
                     autofitRow, autofitCol)
			res + 1
		}
)

setMethod("getBoundingBox", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet, startRow = 0, startCol = 0, endRow = 0, endCol = 0,
             autofitRow = TRUE, autofitCol = TRUE) {
			res <- xlcCall(object, "getBoundingBox", sheet, as.integer(startRow - 1), 
                     as.integer(startCol - 1), as.integer(endRow - 1), as.integer(endCol - 1),
                     autofitRow, autofitCol)
			res + 1
		}
)