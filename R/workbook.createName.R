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
# Creating names in a workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("createName",
	function(object, name, formula, overwrite) standardGeneric("createName"))

setMethod("createName", 
		signature(object = "workbook", name = "character", formula = "character", 
				overwrite = "logical"), 
		function(object, name, formula, overwrite) {
			xlcCall(object@jobj$createName, name, formula, overwrite)
			invisible()
		}
)

setMethod("createName", 
		signature(object = "workbook", name = "character", formula = "character", 
				overwrite = "missing"), 
		function(object, name, formula, overwrite) {
			callGeneric(object, name, formula, FALSE)
		}
)
