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
# Clearing a range from reference
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

setGeneric("clearRangeFromReference",
		function(object, reference) standardGeneric("clearRangeFromReference"))


setMethod("clearRangeFromReference", 
		signature(object = "workbook", reference = "character"), 
		function(object, reference) {
			xlcCall(object, "clearRangeFromReference", reference)
			invisible()
		}
)