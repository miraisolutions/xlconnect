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
# Setting missing values (character or numeric)
#
# Missing values are used when reading and writing data. If data is
# written, the first element of the missing values vector/list is used
# as a missing value identifier. If data is read, values are checked against
# the set of defined missing values - if the value matches
# any of the defined missing values, NA will be assumed.
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setMissingValue",
		function(object, value) standardGeneric("setMissingValue"))

setMethod("setMissingValue", 
		signature(object = "workbook", value = "ANY"), 
		function(object, value) {
		  res = lapply(as.list(value), function(x) {
        if(is.null(x) || is.na(x)) {
          .jnull()
		    } else if(is.character(x)) {
		      .jnew("java/lang/String", x)
		    } else if(is.numeric(x)) {
		      .jnew("java/lang/Double", as.double(x))
		    } else {
		      NULL
		    }
		  })
		  res = res[!sapply(res, is.null)]
		  arr = .jarray(res, contents.class = "java/lang/Object")
		  xlcCall(object, "setMissingValue", arr)
			invisible()
		}
)
