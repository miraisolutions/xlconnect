# Mirai Utility
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

mirai <- NA
class(mirai) <- "mirai"
print.mirai <- function(x, ...) {
	browseURL(url = options("MiraiSolutions.URL")[[1]])
}
