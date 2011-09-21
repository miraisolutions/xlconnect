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
# Setting the missing value string
#
# This defines what string is written to a cell in case a missing value (NA)
# is encountered. If NULL, missing values are represented by blank cells.
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setMissingValue",
		function(object, value) standardGeneric("setMissingValue"))

setMethod("setMissingValue", 
		signature(object = "workbook", value = "ANY"), 
		function(object, value) {
			if(is.null(value))
				xlcCall(object, "setMissingValue", .jarray(.jnull("java/lang/String")))
			else {
				value = unique(as.character(value))
				# convert to array
				res = .jarray(value)
				# NA to Java null
				if(any(is.na(value)))
					res[[which(is.na(value))]] = .jnull("java/lang/String")
				
				xlcCall(object, "setMissingValue", res)
			}
			invisible()
		}
)
