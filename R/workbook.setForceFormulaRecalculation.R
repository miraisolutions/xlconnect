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
# Set a flag to force excel to recalculate formula values
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setForceFormulaRecalculation",
		function(object, sheet, value) standardGeneric("setForceFormulaRecalculation"))

setMethod("setForceFormulaRecalculation", 
		signature(object = "workbook", sheet = "character"), 
		function(object, sheet, value) {
			if(sheet == "*") {
				callGeneric(object, getSheets(object), value)
			} else
				xlcCall(object, "setForceFormulaRecalculation", sheet, value)
			
			invisible()
		}
)

setMethod("setForceFormulaRecalculation", 
		signature(object = "workbook", sheet = "numeric"), 
		function(object, sheet, value) {
			xlcCall(object, "setForceFormulaRecalculation", as.integer(sheet-1), value)
			invisible()
		}
)
