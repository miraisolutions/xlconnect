#############################################################################
#
# XLConnect
# Copyright (C) 2010-2016 Mirai Solutions GmbH
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
# Setting hyperlinks
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setHyperlink",
           function(object, formula, sheet, row, col, type, address) standardGeneric("setHyperlink"))

setMethod("setHyperlink", 
          signature(object = "workbook", formula = "missing", sheet = "numeric"), 
          function(object, formula, sheet, row, col, type, address) {
            xlcCall(object, "setHyperlink", as.integer(sheet - 1), .jseq(row - 1),
                    .jseq(col - 1), type, address)
            invisible()
          }
)

setMethod("setHyperlink", 
          signature(object = "workbook", formula = "missing", sheet = "character"), 
          function(object, formula, sheet, row, col, type, address) {
            xlcCall(object, "setHyperlink", sheet, .jseq(row - 1),
                    .jseq(col - 1), type, address)
            invisible()
          }
)

setMethod("setHyperlink",
          signature(object = "workbook", formula = "character", sheet = "missing"),
          function(object, formula, sheet, row, col, type, address) {
            xlcCall(object, "setHyperlink", formula, type, address)
            invisible()
          }
)
