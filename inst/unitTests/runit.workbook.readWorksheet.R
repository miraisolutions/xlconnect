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
# Tests around reading worksheets from an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.readWorksheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
	
	checkDf <- data.frame(
			"NumericColumn" = c(-23.63, NA, NA, 5.8, 3),
			"StringColumn" = c("Hello", NA, NA, NA, "World"),
			"BooleanColumn" = c(TRUE, FALSE, FALSE, NA, NA),
			"DateTimeColumn" = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07"), tz = "UTC"),
			stringsAsFactors = FALSE
	)
	
	# Read worksheet without specifying range
	# Check that the read data region equals the defined data.frame (*.xls)
	res.index <- readWorksheet(wb.xls, 1)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xls, "Test1")
	checkEquals(res.name, checkDf)
	
	# Read worksheet without specifying range
	# Check that the read data region equals the defined data.frame (*.xlsx)
	res.index <- readWorksheet(wb.xlsx, 1)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xlsx, "Test1")
	checkEquals(res.name, checkDf)
	
	# Read worksheet by specifying a range
	# Check that the read data region equals the defined data.frame (*.xls)
	res.index <- readWorksheet(wb.xls, 2, startRow = 17, startCol = 6, endRow = 22, endCol = 9, header = TRUE)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xls, "Test2", startRow = 17, startCol = 6, endRow = 22, endCol = 9, header = TRUE)
	checkEquals(res.name, checkDf)
	
	# Read worksheet by specifying a range
	# Check that the read data region equals the defined data.frame (*.xlsx)
	res.index <- readWorksheet(wb.xlsx, 2, startRow = 17, startCol = 6, endRow = 22, endCol = 9, header = TRUE)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xlsx, "Test2", startRow = 17, startCol = 6, endRow = 22, endCol = 9, header = TRUE)
	checkEquals(res.name, checkDf)
	
	# Read worksheet by specifying a range via the region argument
	# Check that the read data region equals the defined data.frame (*.xls)
	res.index <- readWorksheet(wb.xls, 2, region = "F17:I22", header = TRUE)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xls, "Test2", region = "F17:I22", header = TRUE)
	checkEquals(res.name, checkDf)
	
	# Read worksheet by specifying a range via the region argument
	# Check that the read data region equals the defined data.frame (*.xlsx)
	res.index <- readWorksheet(wb.xlsx, 2, region = "F17:I22", header = TRUE)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xlsx, "Test2", region = "F17:I22", header = TRUE)
	checkEquals(res.name, checkDf)
	
	# Read worksheet by specifying a range via the region argument (region takes precedence over index specifications)
	# Check that the read data region equals the defined data.frame (*.xls)
	res.index <- readWorksheet(wb.xls, 2, region = "F17:I22", startRow = 88, endCol = 45, header = TRUE)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xls, "Test2", region = "F17:I22", startRow = 88, endCol = 45, header = TRUE)
	checkEquals(res.name, checkDf)
	
	# Read worksheet by specifying a range via the region argument (region takes precedence over index specifications)
	# Check that the read data region equals the defined data.frame (*.xlsx)
	res.index <- readWorksheet(wb.xlsx, 2, region = "F17:I22", startRow = 88, endCol = 45, header = TRUE)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xlsx, "Test2", region = "F17:I22", startRow = 88, endCol = 45, header = TRUE)
	checkEquals(res.name, checkDf)
	
	# Check that an exception is thrown when trying to read from an invalid worksheet (*.xls)
	checkException(readWorksheet(wb.xls, 23))
	checkException(readWorksheet(wb.xls, "SheetDoesNotExist"))
	
	# Check that an exception is thrown when trying to read from an invalid worksheet (*.xlsx)
	checkException(readWorksheet(wb.xlsx, 23))
	checkException(readWorksheet(wb.xlsx, "SheetDoesNotExist"))
	
	# Check that an exception is thrown when attempting to read from a blank sheet (*.xls)
	checkException(readWorksheet(wb.xls, 3))
	checkException(readWorksheet(wb.xls, "Test3"))
	
	# Check that an exception is thrown when attempting to read from a blank sheet (*.xlsx)
	checkException(readWorksheet(wb.xlsx, 3))
	checkException(readWorksheet(wb.xlsx, "Test3"))
	
	checkDf1 <- data.frame(
		A = c(1:2, NA, 3:6, NA),
		B = letters[1:8],
		C = c("z", "y", "x", "w", NA, "v", "u", NA),
		D = c(NA, 1:5, NA, NA),
		stringsAsFactors = FALSE
	)
	
	checkDf2 <- data.frame(
		A = c(rep(NA, 3), 3:6, NA),
		B = c(NA, letters[2:8]),
		C = c("z", "y", "x", "w", NA, "v", "u", NA),
		D = c(NA, 1:5, NA, NA),
		stringsAsFactors = FALSE
	)
	
	# Check that the data bounding box is correctly inferred even if there are blank cells
	# in the last row (*.xls)
	res <- readWorksheet(wb.xls, "Test4")
	checkEquals(res, checkDf1)
	res <- readWorksheet(wb.xls, "Test5")
	checkEquals(res, checkDf2)

	# Check that the data bounding box is correctly inferred even if there are blank cells
	# in the last row (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test4")
	checkEquals(res, checkDf1)
	res <- readWorksheet(wb.xlsx, "Test5")
	checkEquals(res, checkDf2)
	
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
	res <- readWorksheet(wb.xls, sheet = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S")
	checkEquals(res, targetNoForce)
	
	# Check that conversion performs ok (without forcing conversion; *.xlsx)
	res <- readWorksheet(wb.xlsx, sheet = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S")
	checkEquals(res, targetNoForce)
	
	# Check that conversion performs ok (with forcing conversion; *.xls)
	res <- readWorksheet(wb.xls, sheet = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = TRUE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S")
	checkEquals(res, targetForce)
	
	# Check that conversion performs ok (with forcing conversion; *.xlsx)
	res <- readWorksheet(wb.xlsx, sheet = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = TRUE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S")
	checkEquals(res, targetForce)
}
