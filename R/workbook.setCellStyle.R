#############################################################################
#
# XLConnect
# Copyright (C) 2010-2025 Mirai Solutions GmbH
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
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Setting cell styles
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setCellStyle",
		function(object, formula, sheet, row, col, cellstyle) standardGeneric("setCellStyle"))

setMethod("setCellStyle", 
		signature(object = "workbook", formula = "missing", sheet = "numeric"), 
		function(object, formula, sheet, row, col, cellstyle) {
		  xlcCall(object, "setCellStyleSheetIndex", .jseq(as.integer(sheet - 1)), .jsle(row - 1),
		    .jsle(col - 1), .jseq(cellstyle), .recycle = FALSE)
			invisible()
		}
)

setMethod("setCellStyle", 
		signature(object = "workbook", formula = "missing", sheet = "character"), 
		function(object, formula, sheet, row, col, cellstyle) {
		  xlcCall(object, "setCellStyleSheetName", .jseq(sheet), .jsle(row - 1),
		    .jsle(col - 1), .jseq(cellstyle), .recycle = FALSE)
			invisible()
		}
)

setMethod("setCellStyle",
    signature(object = "workbook", formula = "character", sheet = "missing"),
    function(object, formula, sheet, row, col, cellstyle) {
      xlcCall(object, "setCellStyleFormula", formula, cellstyle)
      invisible()
    }
)
