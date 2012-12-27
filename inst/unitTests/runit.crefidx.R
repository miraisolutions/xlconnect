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
# Tests around converting Excel cell references to row & column indices
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.crefidx <- function() {
	
	target <- matrix(c(5, 8, 14, 38), ncol = 2, byrow = TRUE)
	checkEquals(cref2idx(c("$H$5", "$AL$14")), target)
	
	checkEquals(idx2cref(c(5, 8, 14, 38)), c("$H$5", "$AL$14"))
	
	checkEquals(idx2cref(cref2idx(c("$KRE3799", "J$26789", "$DX$357")), absRow = FALSE, absCol = FALSE), c("KRE3799", "J26789", "DX357"))
	
	x <- c(36, 43, 25, 13, 356, 46)
	target <- matrix(x, ncol = 2, byrow = TRUE)
	checkEquals(cref2idx(idx2cref(x)), target)
	
}
