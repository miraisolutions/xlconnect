# Set XLConnect Java Log Level
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

setJavaLogLevel <- function(level) {
	J("com.miraisolutions.xlconnect.utils.Logging")$withLevel(level)
}
