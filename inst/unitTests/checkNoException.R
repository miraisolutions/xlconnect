# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

checkNoException <- function(expr) {
	res <- try(expr)
	checkTrue(!is(res, "try-error"))
}
