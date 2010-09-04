# General XLConnect Settings
# Called by .onLoad which is also passing the package description (pdesc)
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

XLConnectSettings <- function(pdesc) {
	
	# URL to Mirai Solutions GmbH Website
	options(MiraiSolutions.URL = pdesc$URL)
	
	# Date/time format used for conversion to string;
	# This is used for communicating a string-based date/time
	# representation to Java which will then convert it to java.util.Date
	options(XLConnect.dateTimeFormat = "%Y-%m-%d %H:%M:%S")
	
	invisible()
}
