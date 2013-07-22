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
# Utility function to extract row names (stored in a column) of a data.frame
# and to add them as proper row names
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

extractRownames <- function(x, col) {
  if(is(x, "list")) {
    if(is.null(col)) col = list(NULL)
    mapply(extractRownames, x, col, SIMPLIFY = FALSE)
  } else{
  	# use attr(x, "row.names") instead of row.names or rownames
  	# since row.names coerces to character for backward compatibility
  	# see help(row.names) for more information
  	if((is.numeric(col) || is.character(col)) && 
  		!is.null(x[[col]])) {
  
  		attr(x, "row.names") <- if(is.numeric(x[[col]])) as.integer(x[[col]]) else as.character(x[[col]])
  		if(is.numeric(col))
  			x[,-col]
  		else
  			x[, names(x) != col]
  	} else
  		x
  }
}
