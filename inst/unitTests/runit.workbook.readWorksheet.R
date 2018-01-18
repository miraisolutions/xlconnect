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
			"DateTimeColumn" = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
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
  # Test using a negative endRow/endCol
  res.name <- readWorksheet(wb.xls, "Test2", startRow = 17, startCol = 6, endRow = -2, endCol = -1, header = TRUE)
  checkEquals(res.name, checkDf[-nrow(checkDf) + 0:1, -ncol(checkDf)])
	
	# Read worksheet by specifying a range
	# Check that the read data region equals the defined data.frame (*.xlsx)
	res.index <- readWorksheet(wb.xlsx, 2, startRow = 17, startCol = 6, endRow = 22, endCol = 9, header = TRUE)
	checkEquals(res.index, checkDf)
	res.name <- readWorksheet(wb.xlsx, "Test2", startRow = 17, startCol = 6, endRow = 22, endCol = 9, header = TRUE)
	checkEquals(res.name, checkDf)
	# Test using a negative endRow/endCol
	res.name <- readWorksheet(wb.xlsx, "Test2", startRow = 17, startCol = 6, endRow = -2, endCol = -1, header = TRUE)
	checkEquals(res.name, checkDf[-nrow(checkDf) + 0:1, -ncol(checkDf)])
	
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
	
	# Check that reading from a blank sheet returns a data.frame with 0 rows and 0 columns (*.xls)
	checkEquals(readWorksheet(wb.xls, 3), data.frame())
	checkEquals(readWorksheet(wb.xls, "Test3"), data.frame())
	
	# Check that reading from a blank sheet returns a data.frame with 0 rows and 0 columns (*.xlsx)
	checkEquals(readWorksheet(wb.xlsx, 3), data.frame())
	checkEquals(readWorksheet(wb.xlsx, "Test3"), data.frame())
	
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
  # Test with negative endRow/endCol
  res <- readWorksheet(wb.xls, "Test4", endRow = -4, endCol = -2)
  checkEquals(res, checkDf1[-nrow(checkDf1) + 0:3, -ncol(checkDf) + 0:1])
  res <- readWorksheet(wb.xls, "Test5", endRow = -3, endCol = -1)
  checkEquals(res, checkDf2[-nrow(checkDf2) + 0:2, -ncol(checkDf)])

	# Check that the data bounding box is correctly inferred even if there are blank cells
	# in the last row (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test4")
	checkEquals(res, checkDf1)
	res <- readWorksheet(wb.xlsx, "Test5")
	checkEquals(res, checkDf2)
	res <- readWorksheet(wb.xlsx, "Test4", endRow = -4, endCol = -2)
	checkEquals(res, checkDf1[-nrow(checkDf1) + 0:3, -ncol(checkDf) + 0:1])
	res <- readWorksheet(wb.xlsx, "Test5", endRow = -3, endCol = -1)
	checkEquals(res, checkDf2[-nrow(checkDf2) + 0:2, -ncol(checkDf)])
	
	targetNoForce <- data.frame(
			AAA = c(NA, NA, NA, 780.9, NA),
			BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
			CCC = c(TRUE, NA, NA, NA, NA),
			DDD = as.POSIXct(c("1984-01-11 12:00:00", NA, NA, NA, NA)),
			stringsAsFactors = FALSE
	)
	
	targetForce <- data.frame(
			AAA = c(-14.65, NA, 11.70, 780.9, NA),
			BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
			CCC = c(TRUE, TRUE, NA, FALSE, FALSE),
			DDD = as.POSIXct(c("1984-01-11 12:00:00", "2012-02-06 16:15:23", "1984-01-11 12:00:00", 
							NA, "1900-12-22 16:04:48")),
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
	
	target = list(
			AAA = data.frame(
					A = 1:3,
					B = letters[1:3],
					C = c(TRUE, TRUE, FALSE),
					stringsAsFactors = FALSE
			),
			BBB = data.frame(
					D = 4:6,
					E = letters[4:6],
					F = c(FALSE, TRUE, TRUE),
					stringsAsFactors = FALSE
			)
	)
	
	# Check that reading multiple worksheets (by name) returns a named list (*.xls)
	res <- readWorksheet(wb.xls, sheet = c("AAA", "BBB"), header = TRUE)
	checkEquals(res, target)
	
	# Check that reading multiple worksheets (by name) returns a named list (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet = c("AAA", "BBB"), header = TRUE)
	checkEquals(res, target)
	
	target = data.frame(
		"With whitespace" = 1:4,
		"And some other funky characters: _=?^~!$@#%§" = letters[1:4],
		check.names = FALSE,
		stringsAsFactors = FALSE
	)
  
	# Check that reading worksheets with check.names = FALSE works (*.xls)
	res <- readWorksheet(wb.xls, sheet = "VariableNames", header = TRUE, check.names = FALSE)
	checkEquals(res, target)
	
	# Check that reading worksheets with check.names = FALSE works (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet = "VariableNames", header = TRUE, check.names = FALSE)
	checkEquals(res, target)

	
	# Check that attempting to specify both keep and drop throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", header = TRUE, keep=c('A','C'), drop=c('B','D')))
	
	# Check that attempting to specify both keep and drop throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", header = TRUE, keep=c('A','C'), drop=c('B','D')))
	
	# Check that attempting to keep a non-existing column (indicated by header name) throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", header = TRUE, keep=c('A','Z')))
	
	# Check that attempting to keep a non-existing column (indicated by header name) throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", header = TRUE, keep=c('A','Z')))
	
	# Check that attempting to keep a column (indicated by index) out of the bounding box throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", header = TRUE, keep=c(1,5)))
	
	# Check that attempting to keep a column (indicated by index) out of the bounding box throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", header = TRUE, keep=c(1,5)))
	
	# Check that attempting to drop a non-existing column (indicated by header name) throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", header = TRUE, drop=c('A','Z')))
	
	# Check that attempting to drop a non-existing column (indicated by header name) throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", header = TRUE, drop=c('A','Z')))
	
	# Check that attempting to drop a column (indicated by index) out of the bounding box throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", header = TRUE, drop=c(1,5)))
	
	# Check that attempting to drop a column (indicated by index) out of the bounding box throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", header = TRUE, drop=c(1,5)))
	
	checkDfSubset <- data.frame(
			A = c(rep(NA, 3), 3:6, NA),
			C = c("z", "y", "x", "w", NA, "v", "u", NA),
			stringsAsFactors = FALSE
	)
	
	# Check that keeping columns A and C (= by name) works fine (*.xls)
	res <- readWorksheet(wb.xls, "Test5", header = TRUE, keep=c('A','C'))
	checkEquals(res, checkDfSubset)
	
	# Check that keeping columns A and C (= by name) works fine (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test5", header = TRUE, keep=c('A','C'))
	checkEquals(res, checkDfSubset)	
	
	# Check that keeping columns A and C (= by name) with header=FALSE throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, "Test5", header = FALSE, keep=c('A','C')))
	
	# Check that keeping columns A and C (= by name) with header=FALSE throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, "Test5", header = FALSE, keep=c('A','C')))	
	
	# Check that dropping columns B and D (= by name) works fine (*.xls)
	res <- readWorksheet(wb.xls, "Test5", header = TRUE, drop=c('B','D'))
	checkEquals(res, checkDfSubset)
	
	# Check that dropping columns B and D (= by name) works fine (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test5", header = TRUE, drop=c('B','D'))
	checkEquals(res, checkDfSubset)
	
	# Check that dropping columns B and D (= by name) with header=FALSE throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, "Test5", header = FALSE, drop=c('B','D')))
	
	# Check that dropping columns B and D (= by name) with header=FALSE throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, "Test5", header = FALSE, drop=c('B','D')))
	
	# Check that keeping columns 1 and 3 (= by index) works fine (*.xls)
	res <- readWorksheet(wb.xls, "Test5", header = TRUE, keep=c(1,3))
	checkEquals(res, checkDfSubset)
	
	# Check that keeping columns 1 and 3 (= by index) works fine (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test5", header = TRUE, keep=c(1,3))
	checkEquals(res, checkDfSubset)
	
	# Check that dropping columns 2 and 4 (= by index) works fine (*.xls)
	res <- readWorksheet(wb.xls, "Test5", header = TRUE, drop=c(2,4))
	checkEquals(res, checkDfSubset)
	
	# Check that dropping columns 2 and 4 (= by index) works fine (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test5", header = TRUE, drop=c(2,4))
	checkEquals(res, checkDfSubset)
	
	# Go on with checks on a selected area
	# (On a selected area) Check that attempting to specify both keep and drop throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c('B','D'), drop=c('C')))
	
	# (On a selected area) Check that attempting to specify both keep and drop throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c('B','D'), drop=c('C')))
	
	# (On a selected area) Check that attempting to keep a non-existing column (indicated by header name) throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c('B','Z')))
	
	# (On a selected area) Check that attempting to keep a non-existing column (indicated by header name) throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c('B','Z')))
	
	# (On a selected area) Check that attempting to keep a column (indicated by index) out of the bounding box throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c(1,5)))
	
	# (On a selected area) Check that attempting to keep a column (indicated by index) out of the bounding box throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c(1,5)))
	
	# (On a selected area) Check that attempting to drop a non-existing column (indicated by header name) throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, drop=c('B','Z')))
	
	# (On a selected area) Check that attempting to drop a non-existing column (indicated by header name) throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, drop=c('B','Z')))
	
	# (On a selected area) Check that attempting to drop a column (indicated by index) out of the bounding box throws an exception (*.xls)
	checkException(readWorksheet(wb.xls, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, drop=c(1,5)))
	
	# (On a selected area) Check that attempting to drop a column (indicated by index) out of the bounding box throws an exception (*.xlsx)
	checkException(readWorksheet(wb.xlsx, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, drop=c(1,5)))
	
	checkDfAreaSubset <- data.frame(
		B = c(NA, letters[2:7]),
		D = c(NA, 1:5, NA),
		stringsAsFactors = FALSE
	)
	
	# (On a selected area) Check that keeping columns B and D (= by name) works fine (*.xls)
	res <- readWorksheet(wb.xls, "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c('B','D'))
	checkEquals(res, checkDfAreaSubset)
	
	# (On a selected area) Check that keeping columns B and D (= by name) works fine (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c('B','D'))
	checkEquals(res, checkDfAreaSubset)	
	
	# (On a selected area) Check that dropping column C (= by name) works fine (*.xls)
	res <- readWorksheet(wb.xls, "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, drop=c('C'))
	checkEquals(res, checkDfAreaSubset)
	
	# (On a selected area) Check that dropping columns B and D (= by name) works fine (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, drop=c('C'))
	checkEquals(res, checkDfAreaSubset)
	
	# (On a selected area) Check that keeping columns 1 and 3 (= by index) works fine (*.xls)
	res <- readWorksheet(wb.xls, "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c(1,3))
	checkEquals(res, checkDfAreaSubset)
	
	# (On a selected area) Check that keeping columns 1 and 3 (= by index) works fine (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, keep=c(1,3))
	checkEquals(res, checkDfAreaSubset)
	
	# (On a selected area) Check that dropping column 2 (= by index) works fine (*.xls)
	res <- readWorksheet(wb.xls, "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, drop=c(2))
	checkEquals(res, checkDfAreaSubset)
	
	# (On a selected area) Check that dropping column 2 (= by index) works fine (*.xlsx)
	res <- readWorksheet(wb.xlsx, "Test5", startRow = 17, startCol = 7, endRow = 24, endCol=9, header = TRUE, drop=c(2))
	checkEquals(res, checkDfAreaSubset)
	
	# Keeping the same columns from multiple sheets (*.xls)
	res <- readWorksheet(wb.xls, sheet=c("Test1","Test4","Test5"), header = TRUE, keep = c(1,2,3))
	checkEquals(res, list(Test1=checkDf[1:3], Test4=checkDf1[1:3], Test5=checkDf2[1:3]))
	
	# Keeping the same columns from multiple sheets (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet=c("Test1","Test4","Test5"), header = TRUE, keep = c(1,2,3))
	checkEquals(res, list(Test1=checkDf[1:3], Test4=checkDf1[1:3], Test5=checkDf2[1:3]))
	
	# Testing the correct replication of the keep argument (reading from 3 sheets, while keep has length 2) (*.xls)
	res <- readWorksheet(wb.xls, sheet=c("Test1","Test4","Test5"), header = TRUE, keep = list(1,2))
	checkEquals(res, list(Test1=checkDf[1], Test4=checkDf1[2], Test5=checkDf2[1]))	

	# Testing the correct replication of the keep argument (reading from 3 sheets, while keep has length 2) (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet=c("Test1","Test4","Test5"), header = TRUE, keep = list(1,2))
	checkEquals(res, list(Test1=checkDf[1], Test4=checkDf1[2], Test5=checkDf2[1]))	
	
	# Keeping different columns from multiple sheets (*.xls)
	res <- readWorksheet(wb.xls, sheet = c("Test1", "Test4", "Test5"), header = TRUE, keep = list(c(1,2),c(2,3),c(1,3)) )
	checkEquals(res, list(Test1=checkDf[1:2], Test4=checkDf1[2:3], Test5=checkDf2[c(1,3)]))
	
	# Keeping different columns from multiple sheets (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet = c("Test1", "Test4", "Test5"), header = TRUE, keep = list(c(1,2),c(2,3),c(1,3)) )
	checkEquals(res, list(Test1=checkDf[1:2], Test4=checkDf1[2:3], Test5=checkDf2[c(1,3)]))
	
	testAAA = data.frame(
			A = 1:3,
			B = letters[1:3],
			C = c(TRUE, TRUE, FALSE),
			stringsAsFactors = FALSE
	)
	
	#Keeping different columns from multiple sheets (2 keep list elements for 4 sheets) (*.xls)
	res <- readWorksheet(wb.xls, sheet = c("Test1", "Test4", "Test5", "AAA"), header = TRUE, keep = list(c(1,2),c(2,3)) )
	checkEquals(res, list(Test1=checkDf[1:2], Test4=checkDf1[2:3], Test5=checkDf2[1:2], AAA=testAAA[2:3]))
	
	#Keeping different columns from multiple sheets (2 keep list elements for 4 sheets) (*.xls)
	res <- readWorksheet(wb.xlsx, sheet = c("Test1", "Test4", "Test5", "AAA"), header = TRUE, keep = list(c(1,2),c(2,3)) )
	checkEquals(res, list(Test1=checkDf[1:2], Test4=checkDf1[2:3], Test5=checkDf2[1:2], AAA=testAAA[2:3]))
	
  	# Dropping the same columns from multiple sheets (*.xls)
	res <- readWorksheet(wb.xls, sheet=c("Test1","Test4","Test5"), header = TRUE, drop = c(1,2))
	checkEquals(res, list(Test1=checkDf[3:4], Test4=checkDf1[3:4], Test5=checkDf2[3:4]))
	
	# Dropping the same columns from multiple sheets (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet=c("Test1","Test4","Test5"), header = TRUE, drop = c(1,2))
	checkEquals(res, list(Test1=checkDf[3:4], Test4=checkDf1[3:4], Test5=checkDf2[3:4]))
	
	# Testing the correct replication of the drop argument (reading from 3 sheets, while drop has length 2) (*.xls)
	res <- readWorksheet(wb.xls, sheet=c("Test1","Test4","Test5"), header = TRUE, drop = list(1,2))
	checkEquals(res, list(Test1=checkDf[2:4], Test4=checkDf1[c(1,3,4)], Test5=checkDf2[2:4]))

	# Testing the correct replication of the drop argument (reading from 3 sheets, while drop has length 2) (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet=c("Test1","Test4","Test5"), header = TRUE, drop = list(1,2))
	checkEquals(res, list(Test1=checkDf[2:4], Test4=checkDf1[c(1,3,4)], Test5=checkDf2[2:4]))	
	
	# Dropping different columns from multiple sheets (*.xls)
	res <- readWorksheet(wb.xls, sheet = c("Test1", "Test4", "Test5"), header = TRUE, drop = list(c(1,2),c(2,3),c(1,3)) )
	checkEquals(res, list(Test1=checkDf[3:4], Test4=checkDf1[c(1,4)], Test5=checkDf2[c(2,4)]))
	
	# Dropping different columns from multiple sheets (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet = c("Test1", "Test4", "Test5"), header = TRUE, drop = list(c(1,2),c(2,3),c(1,3)) )
	checkEquals(res, list(Test1=checkDf[3:4], Test4=checkDf1[c(1,4)], Test5=checkDf2[c(2,4)]))
	
	# Dropping different columns from multiple sheets (2 drop list elements for 4 sheets) (*.xls)
	res <- readWorksheet(wb.xls, sheet = c("Test1", "Test4", "Test5", "AAA"), header = TRUE, drop = list(c(1,2),c(2,3)) )
	checkEquals(res, list(Test1=checkDf[3:4], Test4=checkDf1[c(1,4)], Test5=checkDf2[3:4], AAA=testAAA[c(1)]))
	
	# Dropping different columns from multiple sheets (2 drop list elements for 4 sheets) (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet = c("Test1", "Test4", "Test5", "AAA"), header = TRUE, drop = list(c(1,2),c(2,3)) )
	checkEquals(res, list(Test1=checkDf[3:4], Test4=checkDf1[c(1,4)], Test5=checkDf2[3:4], AAA=testAAA[c(1)]))

	target1 <- data.frame(
	  Col1 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,  NA, 7, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,  NA), 
	  Col2 = c(NA, NA, NA, 3, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, 13),      
	  Col3 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col4 = c(NA, NA, NA, 4, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, 9, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA),
	  Col5 = c(1, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col6 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col7 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, 10, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col8 = c(2, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col9 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col10 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, 11, NA, NA, NA, NA, NA), 
	  Col11 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col12 = c(NA, NA, NA, NA, NA, NA, 5, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, 12, NA, NA, NA, NA, NA), 
	  Col13 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col14 = c(NA, NA, NA, NA, NA, NA, 6, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col15 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA), 
	  Col16 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, 8, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA)
	)
	
	target2 = data.frame(
	  Col1 = c(9, NA, NA, NA, NA, NA), 
	  Col2 = c(NA,  NA, NA, NA, NA, NA), 
	  Col3 = c(NA, NA, NA, NA, NA, NA), 
	  Col4 = c(10,  NA, NA, NA, NA, NA), 
	  Col5 = c(NA, NA, NA, NA, NA, NA), 
	  Col6 = c(NA,  NA, NA, NA, NA, NA), 
	  Col7 = c(NA, NA, NA, NA, NA, 11)
	)
	
	target3 = data.frame(
	  Col1 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,  NA, NA), 
	  Col2 = c(NA, NA, NA, 9, NA, NA, NA, NA, NA, NA, NA,  NA), 
	  Col3 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA),
	  Col4 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA),     
	  Col5 = c(NA, NA, NA, 10, NA, NA, NA, NA, NA, NA, NA, NA),
	  Col6 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA),
	  Col7 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA),
	  Col8 = c(NA, NA, NA, NA, NA, NA, NA, NA, 11, NA, NA, NA),
	  Col9 = c(NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA)
	)
	
	target4 = as.data.frame(matrix(NA, nrow = 10, ncol = 8))
	names(target4) = paste("Col", 1:8, sep = "")
	
	target5 = data.frame(
	  Col1 = c(NA, NA, NA, NA, 4, NA),
	  Col2 = c(NA, 1, NA, NA, NA, NA)
	)
	
	target6 = data.frame(
	  Col1 = c(NA, NA, NA, NA), 
	  Col2 = c(NA, NA, NA,  4), 
	  Col3 = c(1, NA, NA, NA), 
	  Col4 = c(NA, NA, NA, NA), 
	  Col5 = c(NA,  NA, NA, NA)
	)
	
	target7 = data.frame(
	  Col1 = c(NA, NA, NA, 4),
	  Col2 = c(1, NA, NA, NA)
	)
	
	# Checking bounding-box resolution (*.xls)
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", autofitRow = TRUE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, target1)
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", autofitRow = FALSE, autofitCol = FALSE, header = FALSE)
	checkEquals(res, target1)
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 20, startCol = 5, endRow = 31, endCol = 13,
	                     autofitRow = TRUE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, target2)
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 20, startCol = 5, endRow = 31, endCol = 13,
	                     autofitRow = FALSE, autofitCol = FALSE, header = FALSE)
	checkEquals(res, target3)
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 12, startCol = 5, endRow = 21, endCol = 12,
	                     autofitRow = TRUE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, data.frame())
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 12, startCol = 5, endRow = 21, endCol = 12,
	                     autofitRow = FALSE, autofitCol = FALSE, header = FALSE)
	checkEquals(res, target4)
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9,
	                     autofitRow = FALSE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, target5)
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9,
	                     autofitRow = TRUE, autofitCol = FALSE, header = FALSE)
	checkEquals(res, target6)
	res <- readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9,
	                     autofitRow = TRUE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, target7)
	
	# Checking bounding-box resolution (*.xlsx)
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", autofitRow = TRUE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, target1)
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", autofitRow = FALSE, autofitCol = FALSE, header = FALSE)
	checkEquals(res, target1)
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 20, startCol = 5, endRow = 31, endCol = 13,
	                     autofitRow = TRUE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, target2)
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 20, startCol = 5, endRow = 31, endCol = 13,
	                     autofitRow = FALSE, autofitCol = FALSE, header = FALSE)
	checkEquals(res, target3)
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 12, startCol = 5, endRow = 21, endCol = 12,
	                     autofitRow = TRUE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, data.frame())
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 12, startCol = 5, endRow = 21, endCol = 12,
	                     autofitRow = FALSE, autofitCol = FALSE, header = FALSE)
	checkEquals(res, target4)
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9,
	                     autofitRow = FALSE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, target5)
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9,
	                     autofitRow = TRUE, autofitCol = FALSE, header = FALSE)
	checkEquals(res, target6)
	res <- readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9,
	                     autofitRow = TRUE, autofitCol = TRUE, header = FALSE)
	checkEquals(res, target7)
  
  
	# Cached value tests: Create workbook
	wb.xls <- loadWorkbook(rsrc("resources/testCachedValues.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testCachedValues.xlsx"), create = FALSE)

	# "AllLocal" contains no formulae
	ref.xls.uncached <- readWorksheet(wb.xls, "AllLocal", useCachedValues = FALSE)
	ref.xls.cached <- readWorksheet(wb.xls, "AllLocal", useCachedValues = TRUE)
	# cached and uncached results should be identical
	checkEquals(ref.xls.uncached, ref.xls.cached)

	ref.xlsx.uncached <- readWorksheet(wb.xlsx, "AllLocal", useCachedValues = FALSE)
	ref.xlsx.cached <- readWorksheet(wb.xlsx, "AllLocal", useCachedValues = TRUE)
	checkEquals(ref.xlsx.uncached, ref.xlsx.cached)

	# XLS and XLSX results should be identical
	checkEquals(ref.xlsx.uncached, ref.xls.uncached)

	# the other three named regions reference external worksheets and can't be read
	# with useCachedValues=FALSE
	onErrorCell(wb.xls, XLC$ERROR.STOP)
	checkException(readWorksheet(wb.xls, "HeaderRemote", useCachedValues = FALSE))
	checkException(readWorksheet(wb.xls, "BodyRemote", useCachedValues = FALSE))
	checkException(readWorksheet(wb.xls, "AllRemote", useCachedValues = FALSE))

	onErrorCell(wb.xlsx, XLC$ERROR.STOP)
	checkException(readWorksheet(wb.xlsx, "HeaderRemote", useCachedValues = FALSE))
	checkException(readWorksheet(wb.xlsx, "BodyRemote", useCachedValues = FALSE))
	checkException(readWorksheet(wb.xlsx, "AllRemote", useCachedValues = FALSE))

	res <- readWorksheet(wb.xls, "HeadersRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
	res <- readWorksheet(wb.xls, "BodyRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
	res <- readWorksheet(wb.xls, "BothRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)

	res <- readWorksheet(wb.xlsx, "HeadersRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
	res <- readWorksheet(wb.xlsx, "BodyRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
	res <- readWorksheet(wb.xlsx, "BothRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
  
  # Check that reading cached cell values in conjunction with converting cell values to string
  # does not lead to cell formulas being returned (see github issue #52)
  res <- readWorksheetFromFile(rsrc("resources/testBug52.xlsx"), sheet = 1, useCachedValues = TRUE)
  expected <- data.frame(
    Var1 = c(2, 4, 6),
    Var2 = c("2", "nope", "6"),
    Var3 = c(NA, 4, 6),
    Var4 = c(2, 4, 6),
    stringsAsFactors = FALSE
  )
  checkEquals(expected, res)
  
  # Check that dimensionality is not dropped when reading in a worksheet with rownames = x 
  # (see github issue #49)
  expected = data.frame(B = 1:5, row.names = letters[1:5])
	res <- readWorksheetFromFile(rsrc("resources/testBug49.xlsx"), sheet = 1, rownames = 1)
  checkEquals(expected, res)
  
  # Check that dates are correctly converted to string in 1904-windowing based Excel files
  # (see github issue #53)
  expected = data.frame(
    A = c("2003-04-06", "2014-10-30", "abc"),
    stringsAsFactors = FALSE
  )
  res <- readWorksheetFromFile(rsrc("resources/testBug53.xlsx"), sheet = 1, dateTimeFormat = "%Y-%m-%d")
  checkEquals(expected, res)
  
  # Check that numbers are correctly converted to string 1904-windowing based Excel files
	expected = data.frame(A = as.POSIXct(c("2015-12-01", "2015-11-17", "1984-01-11")))
	res <- readWorksheetFromFile(rsrc("resources/testBug53.xlsx"), sheet = 2, colTypes = "POSIXt", forceConversion = TRUE)
  checkEquals(expected, res)
}
