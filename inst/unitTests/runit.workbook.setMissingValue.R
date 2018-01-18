#############################################################################
#
# XLConnect
# Copyright (C) 2010-2018 Mirai Solutions GmbH
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
# Tests around setting missing value string
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.setMissingValue <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookSetMissingValue.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookSetMissingValue.xlsx"), create = TRUE)
	
	# Test data
	data <- data.frame(A = c(4.2, -3.2, NA, 1.34), B = c("A", NA, "C", "D"), stringsAsFactors = FALSE)
	
	name = "missing"

	createSheet(wb.xls, name = name)
	createName(wb.xls, name = name, formula = paste(name, "$A$1", sep = "!"))
	
	createSheet(wb.xlsx, name = name)
	createName(wb.xlsx, name = name, formula = paste(name, "$A$1", sep = "!"))
	
	# Check that writing and reading the data with the default missing value behaviour returns
	# the original data (*.xls)
	writeNamedRegion(wb.xls, data, name = name)
	res <- readNamedRegion(wb.xls, name = name)
	checkEquals(data, res)
	
	# Check that writing and reading the data with the default missing value behaviour returns
	# the original data (*.xlsx)
	writeNamedRegion(wb.xlsx, data, name = name)
	res <- readNamedRegion(wb.xlsx, name = name)
	checkEquals(data, res)
	
	# Check that writing and reading the data with a specific missing value string 
	# returns the original data but with the numeric column as a character and corresonding
	# missing value (*.xls)
	expect <- data.frame(A = c("4.2", "-3.2", "missing", "1.34"), B = c("A", "missing", "C", "D"), stringsAsFactors = FALSE)
	setMissingValue(wb.xls, value = "missing")
	writeNamedRegion(wb.xls, data, name = name)
	# Reset missing value string such that missing value string is read as string
	setMissingValue(wb.xls, value = NULL)
	res <- readNamedRegion(wb.xls, name = name)
	checkEquals(expect, res)
	
	# Check that writing and reading the data with a specific missing value string 
	# returns the original data but with the numeric column as a character and corresonding
	# missing value (*.xlsx)
	expect <- data.frame(A = c("4.2", "-3.2", "missing", "1.34"), B = c("A", "missing", "C", "D"), stringsAsFactors = FALSE)
	setMissingValue(wb.xlsx, value = "missing")
	writeNamedRegion(wb.xlsx, data, name = name)
	# Reset missing value string such that missing value string is read as string
	setMissingValue(wb.xlsx, value = NULL)
	res <- readNamedRegion(wb.xlsx, name = name)
	checkEquals(expect, res)
	
	# Check that writing and reading the data with a specific missing value string 
	# returns the original data (*.xls)
	setMissingValue(wb.xls, value = "missing")
	writeNamedRegion(wb.xls, data, name = name)
	res <- readNamedRegion(wb.xls, name = name)
	checkEquals(data, res)
	
	# Check that writing and reading the data with a specific missing value string 
	# returns the original data (*.xls)
	setMissingValue(wb.xlsx, value = "missing")
	writeNamedRegion(wb.xlsx, data, name = name)
	res <- readNamedRegion(wb.xlsx, name = name)
	checkEquals(data, res)
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookMissingValue.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookMissingValue.xlsx"), create = FALSE)
	
	expect <- data.frame(
		A = c(NA, -3.2, 3.4, NA, 8, NA),
		B = c("a", NA, "c", "x", "a", "o"),
		C = c(TRUE, TRUE, FALSE, NA, FALSE, NA),
		D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
		stringsAsFactors = FALSE
	)
	
	# Check that reading data with multiple missing value strings works (*.xls)
	setMissingValue(wb.xls, value = c("NA", "missing", "empty"))
	res <- readNamedRegion(wb.xls, name = "Missing1")
	checkEquals(res, expect)
	
	# Check that reading data with multiple missing value strings works (*.xlsx)
	setMissingValue(wb.xlsx, value = c("NA", "missing", "empty"))
	res <- readNamedRegion(wb.xlsx, name = "Missing1")
	checkEquals(res, expect)
  
	expect <- data.frame(
	  A = c(NA, -3.2, NA, NA, 8, NA),
	  B = c("a", NA, "c", "x", "a", "o"),
	  C = c(TRUE, NA, FALSE, NA, FALSE, NA),
	  D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
	  stringsAsFactors = FALSE
	)
  
	# Check that reading data with multiple missing value strings works (*.xls)
	setMissingValue(wb.xls, value = list("NA", "missing", "empty", -9999))
	res <- readNamedRegion(wb.xls, name = "Missing2")
	checkEquals(res, expect)
	
	# Check that reading data with multiple missing value strings works (*.xlsx)
	setMissingValue(wb.xlsx, value = list("NA", "missing", "empty", -9999))
	res <- readNamedRegion(wb.xlsx, name = "Missing2")
	checkEquals(res, expect)
}
