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
# Creates a data.frame from a RDataFrameWrapper jobjRef (rJava) instance.
#
# This is used when reading data from Excel workbooks.
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

dataframeFromJava <- function(df, check.names) {
	jTryCatch({
  	if(!is(df, "jobjRef"))
  		stop("Invalid object - object of class 'jobjRef' required!")
  	
  	columnTypes = .jcall(df, "[S", "getColumnTypes")
  	columnNames = .jcall(df, "[S", "getColumnNames")
  	
  	# Init result list to contain column vectors
  	res = list()
  	for(i in seq(along = columnTypes)) {
  		# Note, Java indices are 0-based while R's are 1-based ...
  		jIndex = as.integer(i - 1)
  		
  		v = switch(columnTypes[i],
  				
  				"Numeric" = {
  					as.vector(.jcall(df, "[D", "getNumericColumn", jIndex))
  				},
  				
  				"String" = {
  					as.vector(.jcall(df, "[S", "getStringColumn", jIndex))
  				},
  				
  				"Boolean" = {
  					as.vector(.jcall(df, "[Z", "getBooleanColumn", jIndex))
  				},
  				
  				"DateTime" = {
            d = as.POSIXct("1970-01-01", tz = "UTC") + (as.vector(.jcall(df, "[J", "getDateTimeColumn", jIndex)) / 1000)
            attr(d, "tzone") = ""
            d
  				},
  				
  				stop("Unsupported column type detected!")
  		)
  		
  		# Put missings back in place;
  		# note that Java primitives are communicated back which 
  		# don't support any missings - therefore this step
  		v[.jcall(df, "[Z", "isMissing", jIndex)] = NA
      res[[i]] = v
  	}
  	
  	# Apply names
  	names(res) = columnNames
  	
  	data.frame(res, check.names = check.names, stringsAsFactors = FALSE)
	})
}
