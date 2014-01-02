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
# "Normalizes" a data.frame with respect to column types.
# This is necessary for comparing data.frames written to Excel and then
# read back in from Excel.
#
# Note that Excel does not know e.g. factor types. Factor variables
# in fact are written to Excel as ordinary strings. Therefore, when read
# back in to R, they are treated as character variables.
#
# 'normalizeDataframe' is used for RUnit tests to compare data.frame's
# written to and read from Excel
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

normalizeDataframe <- function(df, replaceInf = FALSE) {
	att = attr(df, "row.names")
	res <- lapply(df, 
		function(col) {
      if(is(col, "numeric")) {
        if(replaceInf) {
          col[is.infinite(col)] = NA # there are no infinites in Excel
          col
        } else {
          col
        }
			} else if(is(col, "logical") || is(col, "character")) {
        col
			} else if(is(col, "Date") || is(col, "POSIXt")) {
        ms = round(as.numeric(as.POSIXct(col)), 3) # only consider up to milliseconds
        d = as.POSIXct("1970-01-01", tz = "UTC") + ms
        attr(d, "tzone") = ""
        d
			} else
				as.character(col)
		}
	)
	res = data.frame(res, stringsAsFactors = F)
	attr(res, "row.names") = att
	res
}
