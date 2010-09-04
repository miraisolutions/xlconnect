# XLConnect Unit Testing Framework
# See also http://rwiki.sciviews.org/doku.php?id=developers:runit
#
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

# RUnit is required for unit testing
if(require("RUnit", quietly = TRUE)) {

	pkg <- "XLConnect"

	# Load library built by R CMD check
	library(package = pkg, character.only = TRUE)
	
	# Load the namespace to allow testing of private functions
	if(is.element(pkg, loadedNamespaces())) {
		attach(loadNamespace(pkg), 
			name = paste("namespace", pkg, sep = ":"), pos = 3)
	}
	
	# Source additional files needed by testing framework
	source("normalizeDataframe.R")
	
	# Set up and run test suite
	TestSuite <- defineTestSuite(paste(pkg, "Test Suite"), dirs = ".")
	TestResult <- runTestSuite(TestSuite)
	
	# Print test results
	protocol <- paste(pkg, "Unit Testing")
	txtProtocol <- paste(protocol, ".txt", sep = "")
	htmlProtocol <- paste(protocol, ".html", sep = "")
	printTextProtocol(TestResult, fileName = txtProtocol)
	printHTMLProtocol(TestResult, fileName = htmlProtocol)
	
	# Show HTML Test Protocol
	browseURL(url = file.path(getwd(), htmlProtocol))
	
} else {
	warning("Cannot run Unit Tests -- Package RUnit is not available!")
}
