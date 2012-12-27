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
# Wraps an argument in a list if it not already is a list.
# Further extracts the 'jobj' slot from S4 classes (such as 'workbook'
# and 'cellstyle').
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

wrapList <- function(x) {
	if(!is(x, "list")) x <- list(x)
	# Extract 'jobj' slot for S4 classes (such as 'workbook' & 'cellstyle')
	x <- lapply(x, function(y) {
		if(!is(y, "jobjRef") && "jobj" %in% slotNames(y)) y@jobj
		else y
	})
	x
}
