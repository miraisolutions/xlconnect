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
# Converting row and column based area references to Excel area references
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

idx2aref <- function(x) {
	if(!is.numeric(x)) stop("x must be a numeric matrix or vector of indices!")
	if(!is.matrix(x)) x <- matrix(x, ncol = 4, byrow = TRUE)
	apply(x, 1, function(xx) {
		paste(
			idx2cref(xx[1:2], absRow = FALSE, absCol = FALSE),
			idx2cref(xx[3:4], absRow = FALSE, absCol = FALSE),
			sep = ":")
	})
}
