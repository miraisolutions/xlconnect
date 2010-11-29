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
# Creates a Java RDataFrameWrapper instance from a data.frame (non-data.frame 
# objects are coerced to data.frame's). The resulting jobjRef instance (rJava)
# can be used when calling Java methods.
#
# This is used when writing data.frame's to Excel workbooks.
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

dataframeToJava <- function(df) {
	# Force data.frame
	if(!is.data.frame(df)) {
		warning("Supplied object is not a data.frame! Trying to continue by converting it to a data.frame.")
		df <- as.data.frame(df)
	}
	
	dFrame <- jTryCatch(new(J("com.miraisolutions.xlconnect.integration.r.RDataFrameWrapper")))
	cnames <- colnames(df)
	for(i in seq(along = df)) {
		v <- df[, i]
		
		if(is(v, "numeric")) {
			jTryCatch(dFrame$addNumericColumn(cnames[i], .jarray(as.double(v)), .jarray(is.na(v))))
		}
		else if(is(v, "logical")) {
			jTryCatch(dFrame$addBooleanColumn(cnames[i], .jarray(v), .jarray(is.na(v))))
		}
		else if(is(v, "character")) {
			jTryCatch(dFrame$addStringColumn(cnames[i], .jarray(v), .jarray(is.na(v))))
		}
		else if(is(v, "factor")) {
			v <- as.character(v)
			jTryCatch(dFrame$addStringColumn(cnames[i], .jarray(v), .jarray(is.na(v))))
		}
		else if(is(v, "Date") || is(v, "POSIXt")) {
			v <- format(v, format = options("XLConnect.dateTimeFormat")[[1]])
			jTryCatch(dFrame$addDateTimeColumn(cnames[i], .jarray(v), .jarray(is.na(v))))
		}
		else
			stop("Unsupported data type (class) detected!")			
	}
	
	dFrame
}
