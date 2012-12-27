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
# Converting cell references to row and column indices
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

cref2idx <- function(x) {
	if(!is.character(x)) stop("x must be a vector of cell references (character)!")
	t(sapply(x, function(xx) {
						cf <- new(J("org.apache.poi.ss.util.CellReference"), xx)
						c(cf$getRow(), cf$getCol())
					}, USE.NAMES = FALSE) + 1)
}
