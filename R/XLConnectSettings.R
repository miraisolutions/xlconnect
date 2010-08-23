# General XLConnect Settings
# Called by .First.lib
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

XLConnectSettings <- function() {
	
	# Date/time format used for conversion to string;
	# This is used by 'prepareForXLConnect' for communicating a string-based
	# date/time representation to Java which will then convert it to java.util.Date
	options(XLConnect.dateTimeFormat = "%Y-%m-%d %H:%M:%S")
	
	invisible()
}
