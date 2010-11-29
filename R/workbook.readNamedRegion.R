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
# Reading named regions from an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("readNamedRegion",
	function(object, name, header) standardGeneric("readNamedRegion"))

setMethod("readNamedRegion", 
	signature(object = "workbook", name = "character", header = "logical"), 
	function(object, name, header) {	
		# returns a list of RDataFrameWrapper Java object references
		dataFrame <- xlcCall(object@jobj$readNamedRegion, name, header, SIMPLIFY = FALSE)
		# construct data.frame
		dataFrame <- lapply(dataFrame, dataframeFromJava)
		names(dataFrame) <- name
		
		# Return data.frame directly in case only one data.frame is read
		if(length(dataFrame) == 1) dataFrame[[1]]
		else dataFrame
	}
)

setMethod("readNamedRegion", 
		signature(object = "workbook", name = "character", header = "missing"), 
		function(object, name, header) {	
			callGeneric(object, name, TRUE)
		}
)
