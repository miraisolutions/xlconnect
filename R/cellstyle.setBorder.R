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
# Specifying borders for cell styles
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setBorder",
		function(object, side, type, color) standardGeneric("setBorder"))

setMethod("setBorder", 
		signature(object = "cellstyle"), 
		function(object, side, type, color) {
			side <- tolower(side)
			if("all" %in% side)
				side <- c("bottom", "left", "right", "top")
			
			type <- as.integer(rep(type, length = length(side)))
			color <- as.integer(rep(color, length = length(side)))
			
			jTryCatch(object@jobj$setBorder(.jarray(side), .jarray(type), .jarray(color)))
		}
)
