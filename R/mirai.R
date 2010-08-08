# Mirai Utility
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

mirai <- "http://www.mirai-solutions.com"
class(mirai) <- "mirai"
print.mirai <- function(x, ...) {
	browseURL(url = mirai)
}
