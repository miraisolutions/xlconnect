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
# Retrieving a flag that forces Excel to recalculate formulae on reload
# 
# Author: Thomas Themel, Mirai Solutions GmbH
#
#############################################################################

setGeneric("getForceFormulaRecalculation",
	function(object, sheet) standardGeneric("getForceFormulaRecalculation"))

setMethod("getForceFormulaRecalculation", 
	signature(object = "workbook", sheet = "numeric"), 
	function(object, sheet) {
		xlcCall(object, "getForceFormulaRecalculation", as.integer(sheet-1))
	}
)

setMethod("getForceFormulaRecalculation", 
	signature(object = "workbook", sheet = "character"), 
	function(object, sheet) {
		if(sheet == "*")
			callGeneric(object, getSheets(object))
		else
			xlcCall(object, "getForceFormulaRecalculation", sheet)
	}
)
