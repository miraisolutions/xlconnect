#############################################################################
#
# XLConnect
# Copyright (C) 2010-2021 Mirai Solutions GmbH
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
# Querying available Excel tables in a workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("getTables",
	function(object, sheet, simplify = TRUE) standardGeneric("getTables"))

setMethod("getTables", 
	signature(object = "workbook", sheet = "numeric"), 
	function(object, sheet, simplify = TRUE) {
	  res <- xlcCall(object, "getTables", as.integer(sheet - 1), .simplify = FALSE)
    if(simplify && (length(res) == 1)) {
      res[[1]]
    } else {
      res
    }
	}
)

setMethod("getTables", 
  signature(object = "workbook", sheet = "character"), 
  function(object, sheet, simplify = TRUE) {
    res <- xlcCall(object, "getTables", sheet, .simplify = FALSE)
    if(simplify && (length(res) == 1)) {
      res[[1]]
    } else {
      res
    }
  }
)
