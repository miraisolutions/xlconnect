# XLConnect Unit Testing Framework
# Reference: http://rwiki.sciviews.org/doku.php?id=developers:runit
#
# Adapted by: Martin Studer, Mirai Solutions GmbH
###############################################################################

# RUnit is required for unit testing
if(require("RUnit", quietly = TRUE)) {

	pkg <- "XLConnect"
	if(Sys.getenv("RCMDCHECK") == "FALSE") {
		## Path to unit tests for standalone running under Makefile (not R CMD check)
		## PKG/tests/../inst/unitTests
		path <- file.path(getwd(), "..", "inst", "unitTests")
	} else {
		## Path to unit tests for R CMD check
		## PKG.Rcheck/tests/../PKG/unitTests
		path <- system.file(package=pkg, "unitTests")
	}
	cat("\nRunning Unit Tests\n")
	print(list(pkg=pkg, getwd=getwd(), pathToUnitTests=path))
	
	# Function to be used by unit tests to refer to resources
	# in unitTests folder
	rsrc <- function(resource) {
		file.path(path, resource)
	}
	
	# Load library built by R CMD check
	library(package = pkg, character.only = TRUE)
	
	# Load the namespace to allow testing of private functions
	if(is.element(pkg, loadedNamespaces())) {
		attach(loadNamespace(pkg), 
			name = paste("namespace", pkg, sep = ":"), pos = 3)
	}
	
	# Source additional files needed by testing framework
	source(file.path(path, "checkNoException.R"))
	source(file.path(path, "normalizeDataframe.R"))
	
	# Set up and run test suite
	TestSuite <- defineTestSuite(paste(pkg, "Test Suite"), dirs = path)
	TestResult <- runTestSuite(TestSuite)
	
	# Print test results
	protocol <- paste(pkg, "Unit Testing")
	txtProtocol <- file.path(path, paste(protocol, ".txt", sep = ""))
	htmlProtocol <- file.path(path, paste(protocol, ".html", sep = ""))
	printTextProtocol(TestResult, fileName = txtProtocol)
	printHTMLProtocol(TestResult, fileName = htmlProtocol)
	
	# Show HTML Test Protocol
	browseURL(url = htmlProtocol)
	
	## Return stop() to cause R CMD check stop in case of
	##  - failures i.e. FALSE to unit tests or
	##  - errors i.e. R errors
	tmp <- getErrors(TestResult)
	if(tmp$nFail > 0 | tmp$nErr > 0) {
		stop(paste("\n\nUnit Testing failed (#test failures: ", tmp$nFail,
						", #R errors: ",  tmp$nErr, ")\n\n", sep=""))
	}
} else {
	warning("Cannot run Unit Tests -- Package RUnit is not available!")
}
