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
# Setting sheet color
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setSheetColor",
		function(object, sheet, color) standardGeneric("setSheetColor"))

setMethod("setSheetColor", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, color) {
			xlcCall(object, "setSheetColor", as.integer(sheet - 1), as.integer(color))
			invisible()
		}
)

setMethod("setSheetColor", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet, color) {
			xlcCall(object, "setSheetColor", sheet, as.integer(color))
			invisible()
		}
)
