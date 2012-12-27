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
# Writing large datasets to Excel
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

# Increase Java's maximum heap space;
# this option must be set before the underlying JVM is initialized
# and therefore MUST happen before XLConnect is loaded
options(java.parameters = "-Xmx1024m")

if(any(is.element(c("package:XLConnect", "package:rJava"), search()))) {
	msg <- paste(
		"XLConnect and/or rJava are already attached.",
		"As such, a JVM instance may already be running.",
		"Please restart R and set the Java parameters before XLConnect",
		"or any packages that start a JVM instance are loaded.",
		sep = "\n"
	)
	warning(msg)
}

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "large.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Create a large dummy data.frame
set.seed(292)
n <- 50000
dfLarge <- data.frame(
	A = rnorm(n),
	B = sample(letters, size = n, replace = TRUE),
	C = rnorm(n) > 0,
	D = rep(as.Date("2010-09-17"), n),
	E = rnorm(n),
	F = sample(LETTERS, size = n, replace = TRUE),
	G = 1:n
)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Create a worksheet called 'Dummy'
createSheet(wb, name = "Dummy")

# Write large data.frame to worksheet 'Dummy' created above
writeWorksheet(wb, dfLarge, sheet = "Dummy")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") browseURL(file.path(getwd(), demoExcelFile))
}
