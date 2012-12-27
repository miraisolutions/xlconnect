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
# Setting worksheet position
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setSheetPos",
		function(object, sheet, pos) standardGeneric("setSheetPos"))

setMethod("setSheetPos", 
		signature(object = "workbook", sheet = "character", pos = "numeric"), 
		function(object, sheet, pos) {
			xlcCall(object, "setSheetPos", sheet, as.integer(pos - 1))
			invisible()
		}
)

setMethod("setSheetPos", 
		signature(object = "workbook", sheet = "character", pos = "missing"), 
		function(object, sheet, pos) {
			callGeneric(object, sheet, seq(along = sheet) - 1)
		}
)
