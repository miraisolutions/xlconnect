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
# Writing named regions to an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("writeNamedRegion",
	function(object, data, name, header = TRUE, rownames = NULL) standardGeneric("writeNamedRegion"))

setMethod("writeNamedRegion", 
	signature(object = "workbook", data = "ANY"), 
	function(object, data, name, header = TRUE, rownames = NULL) {
		data <- includeRownames(data, rownames)
		# pass data.frame's to Java - construct RDataFrameWrapper Java object references
		data <- lapply(wrapList(data), dataframeToJava)
		xlcCall(object, "writeNamedRegion", data, name, header, SIMPLIFY = FALSE)
		invisible()
	}
)
