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
# Tests around converting Excel columns to and from indices
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.colidx <- function() {
	
	checkEquals(col2idx(c("A", "AA", "HZR")), c(1, 27, 6102))
	
	checkEquals(idx2col(c(1, 27, 6102)), c("A", "AA", "HZR"))
	
	checkEquals(idx2col(col2idx(c("AWT", "FRT"))), c("AWT", "FRT"))
	
	checkEquals(col2idx(idx2col(c(3628, 867))), c(3628, 867))
	
	checkEquals(idx2col(c(0, -1, -2628)), rep("", 3))
}
