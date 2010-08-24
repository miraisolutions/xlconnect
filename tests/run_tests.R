# XLConnect Unit Testing
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(RUnit)

# Set up and run test suite
XLConnectTestSuite <- defineTestSuite("XLConnect Test Suite", dirs = ".")
XLConnectTestResult <- runTestSuite(XLConnectTestSuite)

# Print test results
printTextProtocol(XLConnectTestResult, fileName = "XLConnect Unit Testing.txt")
printHTMLProtocol(XLConnectTestResult, fileName = "XLConnect Unit Testing.html")
