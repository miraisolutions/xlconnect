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
# Accessing named cell styles
# 
# Author: Thomas Themel, Mirai Solutions GmbH
#
#############################################################################

setGeneric("getCellStyle",
		function(object, name) standardGeneric("getCellStyle"))

setMethod("getCellStyle", 
		signature(object = "workbook"), 
		function(object, name) {
			jobj <- jTryCatch(object@jobj$getCellStyle(name))
                        new("cellstyle", jobj = jobj)
		}
)
