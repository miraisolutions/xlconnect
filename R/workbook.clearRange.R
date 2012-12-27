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
# Clearing a range
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

setGeneric("clearRange",
		function(object, sheet, coords) standardGeneric("clearRange"))

setMethod("clearRange", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, coords) {
			if(!is.numeric(coords))
				stop("Argument coords needs to be numeric")
				
			if(!is.matrix(coords))
				coords = matrix(coords, ncol = 4, byrow = TRUE)
			
			coords = lapply(split(as.integer(coords - 1), row(coords)), .jarray)
			
			xlcCall(object, "clearRange", as.integer(sheet - 1), coords)
			invisible()
		}
)

setMethod("clearRange", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet, coords) {
			if(!is.numeric(coords))
				stop("Argument coords needs to be numeric")
			
			if(!is.matrix(coords))
				coords = matrix(coords, ncol = 4, byrow = TRUE)
			
			coords = lapply(split(as.integer(coords - 1), row(coords)), .jarray)
			
			xlcCall(object, "clearRange", sheet, coords)
			invisible()
		}
)