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
# RUnit test-runner script
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

# Set timezone to UTC
Sys.setenv("TZ" = "UTC")

# Limit number of GC threads; set timezone
options(java.parameters = c("-XX:+UseParallelGC", "-XX:ParallelGCThreads=1", paste0("-Duser.timezone=", Sys.timezone())))

# Load library built by R CMD check
library(package = "XLConnect", character.only = TRUE)
require(rJava)

rsrc <- function(resource) {
  file.path(options()$path.unit.tests, resource)
}

checkNoException <- function(expr) {
        res <- try(expr)
        checkTrue(!is(res, "try-error"))
}

runUnitTests <- function() {
	
	pkg <- "XLConnect"
	
	# Option to determine if full test suite should be run
	options("FULL.TEST.SUITE" = Sys.getenv("FULL_TEST_SUITE") == "1")

	# RUnit is required for unit testing
	if(require("RUnit", quietly = TRUE)) {
		if(Sys.getenv("RCMDCHECK") == "FALSE") {
			# Path to unit tests for standalone running under Makefile (not R CMD check)
			path <- file.path(getwd(), "..", "inst", "unitTests")
		} else {
			# Path to unit tests for 'R CMD check' and as part of public API
			path <- system.file(package = pkg, "unitTests")
		}
		cat("\nRunning Unit Tests\n")
		print(list(WorkingDir = getwd(), PathToUnitTests = path))
		
		# Add path to unit tests as option
		# (used by rsrc function in unit tests)
		options(path.unit.tests = path)
		
		# Set up and run test suite
		Sys.setlocale(category = "LC_NUMERIC", locale = "C")
		jlocale = J("java.util.Locale")
		jlocale$setDefault(jlocale$US)
		orig.opts <- options(encoding = "UTF-8")
		TestSuite <- defineTestSuite(paste(pkg, "Test Suite"), dirs = path)
		TestResult <- runTestSuite(TestSuite)
		options(orig.opts)
		
		# Test protocol files
		protocol <- file.path(getwd(), paste(pkg, "Unit_Tests", sep = "_"))
		txtProtocol <- paste(protocol, ".txt", sep = "")
		htmlProtocol <- paste(protocol, ".html", sep = "")
		
		# Print (summary) test protocol to stdout
		printTextProtocol(TestResult, showDetails = FALSE)
		# Write detailed test protocol to text file
		printTextProtocol(TestResult, showDetails = TRUE, fileName = txtProtocol)
		# Write HTML protocol
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
		warning("Cannot run unit tests -- Package 'RUnit' is not available!")
	}
	
	invisible()
}

# Run unit tests
runUnitTests()

