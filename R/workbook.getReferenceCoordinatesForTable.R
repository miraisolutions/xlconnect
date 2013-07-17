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
# Querying coordinates of Excel tables (Office 2007+)
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("getReferenceCoordinatesForTable",
		function(object, sheet, table) standardGeneric("getReferenceCoordinatesForTable"))

setMethod("getReferenceCoordinatesForTable", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, table) {
			res <- xlcCall(object, "getReferenceCoordinatesForTable", as.integer(sheet - 1), table)
      if(is.numeric(res)) { matrix(res, nrow = 2, byrow = TRUE) + 1 } else { res } 
		}
)

setMethod("getReferenceCoordinatesForTable", 
          signature(object = "workbook", sheet = "character"), 
          function(object, sheet, table) {
            res <- xlcCall(object, "getReferenceCoordinatesForTable", sheet, table)
            if(is.numeric(res)) { matrix(res, nrow = 2, byrow = TRUE) + 1 } else { res } 
          }
)
