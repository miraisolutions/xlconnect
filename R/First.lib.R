# XLConnect package initialization
# Loads rJava to be able to call underlying Java code
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


.First.lib <- function(libname, pkgname) {
	# Print package information
	pdesc <- packageDescription(pkgname)
	cat(pdesc$Package, pdesc$Version, "by", pdesc$Maintainer, "\n", sep = " ")
	cat(pdesc$URL, "\n")

	# Load Java dependencies (all jars inside the java subfolder)
	.jpackage(name = pkgname, jars = "*")
	
	# Perform general XLConnect settings - pass package description
	XLConnectSettings(pdesc)
}
