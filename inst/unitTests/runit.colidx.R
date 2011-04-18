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
# Tests around converting Excel columns to and from indices
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.colidx <- function() {
	
	checkEquals(col2idx("A"), 1)
	checkEquals(col2idx("AA"), 27)
	checkEquals(col2idx("HZR"), 6102)
	
	checkEquals(idx2col(1), "A")
	checkEquals(idx2col(27), "AA")
	checkEquals(idx2col(6102), "HZR")
	
	checkEquals(idx2col(col2idx("AWT")), "AWT")
	checkEquals(idx2col(col2idx("FRT")), "FRT")
	
	checkEquals(col2idx(idx2col(3628)), 3628)
	checkEquals(col2idx(idx2col(867)), 867)
	
	checkEquals(idx2col(0), "")
	checkEquals(idx2col(-1), "")
	checkEquals(idx2col(-2628), "")
}
