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
# Utility function for vectorizing argument lists and default Java exception
# handling (jTryCatch)
#
# Atomic objects are replicated as is, others are wrapped in a list as defined
# by wrapList and then replicated
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

xlcCall <- function(fun, ..., SIMPLIFY = TRUE) {
	args <- list(...)
	args <- lapply(args, function(x) {
		if(is.atomic(x)) x
		else wrapList(x)
	})
	jTryCatch(do.call("mapply", args = c(FUN = fun, args, SIMPLIFY = SIMPLIFY)))
}
