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
# Workbook operators
#
# $: 	Execute workbook methods in object$method(...) form
# [:	Alias for readWorksheet
# [<-:	Alias for writeWorksheet
# [[:	Alias for readNamedRegion
# [[<-:	Alias for writeNamedRegion
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setMethod("$", 
	signature(x = "workbook"),
	function(x, name) {
		g <- getGeneric(name)
		if(is.null(g)) stop("Method undefined for class 'workbook'")
		function(...) g(x, ...)
	}
)

setMethod("[",
	signature(x = "workbook"),
	function(x, i, j, ..., drop = FALSE) {
		readWorksheet(x, sheet = i, ...)
	}
)

setMethod("[<-",
	signature(x = "workbook"),
	function(x, i, j, ..., value) {
		idx = !(i %in% getSheets(x))
		if(any(idx))
			createSheet(x, name = i[idx])
		writeWorksheet(x, data = value, sheet = i, ...)
		x
	}
)

setMethod("[[",
	signature(x = "workbook"),
	function(x, i, j, ...) {
		readNamedRegion(x, name = i, ...)
	}
)

setMethod("[[<-",
	signature(x = "workbook"),
	function(x, i, j, ..., value) {
		idxn = !(i %in% getDefinedNames(x))
		if(any(idxn)) {
			sheet = extractSheetName(j[idxn])
			idxs = !(sheet %in% getSheets(x))
			if(any(idxs))
				createSheet(x, name = sheet[idxs])
			createName(x, name = i[idxn], formula = j[idxn])
		}
		writeNamedRegion(x, data = value, name = i, ...)
		x
	}
)
