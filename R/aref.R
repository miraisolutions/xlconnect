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
# Creates an Excel area reference
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

aref <- function(topLeft, dimension) {
	if(is.character(topLeft))
		topLeft <- as.vector(cref2idx(topLeft))
	if(!is.numeric(topLeft))
		stop("topLeft must be either a cell reference (character) in the form 'A1' or a numeric vector of length 2!")
	if(!is.numeric(dimension) || length(dimension) != 2)
		stop("dimension must be a numeric vector of length 2!")
	
	idx2aref(c(topLeft, topLeft + dimension - 1))
}
