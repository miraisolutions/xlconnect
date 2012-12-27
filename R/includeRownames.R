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
# Utility function to include rownames in a data.frame
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

includeRownames <- function(x, colname) {
  if(is(x, "list")) {
    if(is.null(colname)) colname = list(NULL)
    mapply(includeRownames, x, colname, SIMPLIFY = FALSE)
  } else {
    # Force data.frame
    if(!is.data.frame(x))
      x <- as.data.frame(x)
    
  	# use attr(x, "row.names") instead of row.names or rownames
  	# since row.names coerces to character for backward compatibility
  	# see help(row.names) for more information
  	if(is.character(colname) && !is.null(attr(x, "row.names"))) {
  		res <- cbind(attr(x, "row.names"), x, stringsAsFactors = FALSE)
  		row.names(res) <- NULL
  		names(res)[1] <- colname
  		res
  	} else
  		x
  }
}
