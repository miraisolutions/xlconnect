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
# Wrapper to tryCatch for standard Java exception handling
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

jTryCatch <- function(...) {
	tryCatch(..., Throwable = 
		function(e) {
			if(!is.jnull(e$jobj)) {
				sw <- new(J("java.io.StringWriter"))
				pw <- new(J("java.io.PrintWriter"), sw)
				e$jobj$printStackTrace(pw)
				options("java.stacktrace"=sw$toString())
				stop(paste(class(e)[1], e$jobj$getMessage(), sep = " (Java): "),
					call. = FALSE)
			} else 
				stop("Undefined error occurred!")
		}
	)
}
