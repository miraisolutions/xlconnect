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
# Tests around converting Excel area references to row & column based
# area references
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.arefidx <- function() {
	
	target <- matrix(c(1, 1, 8, 2, 9, 4, 26, 11), ncol = 4, byrow = TRUE)
	checkEquals(aref2idx(c("A1:B8", "D9:K26")), target)
	
	checkEquals(idx2aref(c(3, 2, 7, 9, 8, 5, 48, 17)), c("B3:I7", "E8:Q48"))
	
	checkEquals(idx2aref(aref2idx(c("D27:J54", "AA23:CD129"))), c("D27:J54", "AA23:CD129"))
	
	x <- c(31, 6, 56, 8, 129, 17, 488, 37)
	target <- matrix(x, ncol = 4, byrow = TRUE)
	checkEquals(aref2idx(idx2aref(x)), target)
	
	checkEquals(aref("BB35", c(678, 25)), "BB35:BZ712")
	checkEquals(aref(c(18, 46), c(16, 18)), "AT18:BK33")
}
