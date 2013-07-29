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
  jTryCatch({
  	# Force data.frame
  	if(!is.data.frame(df))
  		df = as.data.frame(df)
  	
  	dFrame = new(J("com.miraisolutions.xlconnect.integration.r.RDataFrameWrapper"))
  	cnames = colnames(df)
  	for(i in seq(along = df)) {
  		v = df[, i]
  		
  		if(is(v, "numeric")) {
  			.jcall(dFrame, "V", "addNumericColumn", cnames[i], .jarray(as.double(v)), .jarray(is.na(v)))
  		}
  		else if(is(v, "logical")) {
  			.jcall(dFrame, "V", "addBooleanColumn", cnames[i], .jarray(v), .jarray(is.na(v)))
  		}
  		else if(is(v, "character")) {
  			.jcall(dFrame, "V", "addStringColumn", cnames[i], .jarray(v), .jarray(is.na(v)))
  		}
  		else if(is(v, "POSIXt")) {
  			.jcall(dFrame, "V", "addDateTimeColumn", cnames[i], .jarray(.jlong(as.numeric(v) * 1000)), .jarray(is.na(v)))
  		} else if(is(v, "Date")) {
        # Important: dates (without time) are treated as being at midnight UTC
  		  .jcall(dFrame, "V", "addDateTimeColumn", cnames[i], .jarray(.jlong(as.numeric(as.POSIXct(v)) * 1000)), .jarray(is.na(v)))
  		} else { # e.g. factor or any other type
  			v = as.character(v)
  			if(length(v) != nrow(df))
  				stop(sprintf("Conversion to character for column %s failed! Check the class.", cnames[i]))
  			.jcall(dFrame, "V", "addStringColumn", cnames[i], .jarray(v), .jarray(is.na(v)))
  		}
  	}
  	
  	dFrame
  })
}
