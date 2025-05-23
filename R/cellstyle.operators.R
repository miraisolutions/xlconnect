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
# Executing workbook methods in object$method(...) form
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setMethod("$", 
		signature(x = "cellstyle"),
		function(x, name) {
			g <- getGeneric(name)
			if(is.null(g)) stop("Method undefined for class 'cellstyle'")
			function(...) g(x, ...)
		}
)
