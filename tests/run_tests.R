# XLConnect Unit Testing Framework
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(RUnit)

# Load XLConnect library built by R CMD check
library(XLConnect, lib.loc = file.path(getwd(), ".."))

# Source additional files needed by testing framework
source("normalizeDataframe.R")

# Set up and run test suite
XLConnectTestSuite <- defineTestSuite("XLConnect Test Suite", dirs = ".")
XLConnectTestResult <- runTestSuite(XLConnectTestSuite)

# Print test results
txtProtocol <- "XLConnect Unit Testing.txt"
htmlProtocol <- "XLConnect Unit Testing.html"
printTextProtocol(XLConnectTestResult, fileName = txtProtocol)
printHTMLProtocol(XLConnectTestResult, fileName = htmlProtocol)

# Show HTML Test Protocol
browseURL(url = file.path(getwd(), htmlProtocol))
