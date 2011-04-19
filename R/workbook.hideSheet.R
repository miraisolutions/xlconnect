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
# Hiding worksheets in a workbook
#
# There are two variants of hiding a worksheet
# - "normal" hiding: users can unhide the worksheets in Excel
# - very hiding: worksheets can be unhidden programmatically only
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################
 
setGeneric("hideSheet",
	function(object, sheet, veryHidden) standardGeneric("hideSheet"))

setMethod("hideSheet", 
		signature(object = "workbook", sheet = "numeric", veryHidden = "logical"), 
		function(object, sheet, veryHidden) {
			xlcCall(object, "hideSheet", as.integer(sheet - 1), veryHidden)
			invisible()
		}
)

setMethod("hideSheet", 
		signature(object = "workbook", sheet = "numeric", veryHidden = "missing"), 
		function(object, sheet, veryHidden) {
			callGeneric(object, sheet, FALSE)
		}
)

setMethod("hideSheet", 
		signature(object = "workbook", sheet = "character", veryHidden = "logical"), 
		function(object, sheet, veryHidden) {
			xlcCall(object, "hideSheet", sheet, veryHidden)
			invisible()
		}
)

setMethod("hideSheet", 
		signature(object = "workbook", sheet = "character", veryHidden = "missing"), 
		function(object, sheet, veryHidden) {
			callGeneric(object, sheet, FALSE)
		}
)
