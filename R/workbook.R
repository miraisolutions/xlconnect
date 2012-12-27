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
# Workbook class definition with initialization method
# This is XLConnect's main entity
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setClass("workbook", representation(filename = "character", jobj = "jobjRef"))

setMethod("initialize", 
		"workbook", 
		function(.Object, filename, create) {
			.Object@filename <- filename
			.Object@jobj <- jTryCatch(new(J("com.miraisolutions.xlconnect.integration.r.RWorkbookWrapper"), filename, create))
			if(is.jnull(.Object@jobj))
				stop("Could not create workbook instance! Got null reference from Java.")
			.Object
		}
)
