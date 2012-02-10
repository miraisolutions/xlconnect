#############################################################################
#
# XLConnect
# Copyright (C) 2010-2012 Mirai Solutions GmbH
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
# Tests around reading named regions from an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.readNamedRegion <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookReadNamedRegion.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReadNamedRegion.xlsx"), create = FALSE)
	
	checkDf <- data.frame(
			"NumericColumn" = c(-23.63, NA, NA, 5.8, 3),
			"StringColumn" = c("Hello", NA, NA, NA, "World"),
			"BooleanColumn" = c(TRUE, FALSE, FALSE, NA, NA),
			"DateTimeColumn" = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07"), tz = "UTC"),
			stringsAsFactors = F
	)
	
	# Check that the read named region equals the defined data.frame (*.xls)
	res <- readNamedRegion(wb.xls, "Test", header = TRUE)
	checkEquals(res, checkDf)
	
	# Check that the read named region equals the defined data.frame (*.xlsx)
	res <- readNamedRegion(wb.xlsx, "Test", header = TRUE)
	checkEquals(res, checkDf)
	
	# Check that attempting to read a non-existing named region throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "NameThatDoesNotExist"))
	
	# Check that attempting to read a non-existing named region throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "NameThatDoesNotExist"))
	
	targetNoForce <- data.frame(
			AAA = c(NA, NA, NA, 780.9, NA),
			BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
			CCC = c(TRUE, NA, NA, NA, NA),
			DDD = as.POSIXct(c("1984-01-11 12:00:00", NA, NA, NA, NA), tz = "UTC"),
			stringsAsFactors = FALSE
	)
	
	targetForce <- data.frame(
			AAA = c(-14.65, NA, 11.70, 780.9, NA),
			BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
			CCC = c(TRUE, TRUE, NA, FALSE, FALSE),
			DDD = as.POSIXct(c("1984-01-11 12:00:00", "2012-02-06 16:15:23", "1984-01-11 12:00:00", 
					NA, "1900-12-22 16:04:48"), tz = "UTC"),
			stringsAsFactors = FALSE
	)
	
	# Check that conversion performs ok (without forcing conversion; *.xls)
	res <- readNamedRegion(wb.xls, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S")
	checkEquals(res, targetNoForce)
	
	# Check that conversion performs ok (without forcing conversion; *.xlsx)
	res <- readNamedRegion(wb.xlsx, name = "Conversion", header = TRUE,
		colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
		forceConversion = FALSE,
		dateTimeFormat = "%d.%m.%Y %H:%M:%S")
	checkEquals(res, targetNoForce)
	
	# Check that conversion performs ok (with forcing conversion; *.xls)
	res <- readNamedRegion(wb.xls, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = TRUE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S")
	checkEquals(res, targetForce)
	
	# Check that conversion performs ok (with forcing conversion; *.xlsx)
	res <- readNamedRegion(wb.xlsx, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = TRUE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S")
	checkEquals(res, targetForce)
}
