#############################################################################
#
# XLConnect
# Copyright (C) 2010-2012 Mirai Solutions GmbH
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

dataframeFromJava <- function(df) {
	
	if(!is(df, "jobjRef"))
		stop("Invalid object - object of class 'jobjRef' required!")
	
	columnTypes <- jTryCatch(df$getColumnTypes())
	columnNames <- jTryCatch(df$getColumnNames())
	
	# Init result list to contain column vectors
	res <- list()
	for(i in seq(along = columnTypes)) {
		# Note, Java indices are 0-based while R's are 1-based ...
		jIndex <- as.integer(i - 1)
		
		switch(columnTypes[i],
				
				"Numeric" = {
					res[[i]] <- as.vector(jTryCatch(df$getNumericColumn(jIndex)))
				},
				
				"String" = {
					res[[i]] <- as.vector(jTryCatch(df$getStringColumn(jIndex)))
				},
				
				"Boolean" = {
					res[[i]] <- as.vector(jTryCatch(df$getBooleanColumn(jIndex)))
				},
				
				"DateTime" = {
					# Convert date/time strings back to POSIXct with timezone UTC
					res[[i]] <- as.POSIXct(as.vector(jTryCatch(df$getDateTimeColumn(jIndex))), 
							format = options("XLConnect.dateTimeFormat")[[1]], tz = "UTC")
				},
				
				stop("Unsupported column type detected!")
		)
		
		# Put missings back in place;
		# note that Java primitives are communicated back which 
		# don't support any missings - therefore this step
		res[[i]][jTryCatch(df$isMissing(jIndex))] <- NA
	}
	
	# Apply names
	names(res) <- columnNames
	
	data.frame(res, stringsAsFactors = FALSE)
}
