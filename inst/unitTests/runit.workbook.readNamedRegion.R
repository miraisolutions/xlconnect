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
			"DateTimeColumn" = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
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
	
	# Check that reading multiple named regions returns a named list (*.xls)
	res <- readNamedRegion(wb.xls, name = c("AAA", "BBB"), header = TRUE)
	checkEquals(res, target)
	
	# Check that reading multiple named regions returns a named list (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name = c("AAA", "BBB"), header = TRUE)
	checkEquals(res, target)
	
	target = data.frame(
		"With whitespace" = 1:4,
		"And some other funky characters: _=?^~!$@#%§" = letters[1:4],
		check.names = FALSE,
		stringsAsFactors = FALSE
	)
	
	# Check that reading named regions with check.names = FALSE works (*.xls)
	res <- readNamedRegion(wb.xls, name = "VariableNames", header = TRUE, check.names = FALSE)
	checkEquals(res, target)
	
	# Check that reading named regions with check.names = FALSE works (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name = "VariableNames", header = TRUE, check.names = FALSE)
	checkEquals(res, target)
	
	# Check that attempting to specify both keep and drop throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "Conversion", header = TRUE, keep=c('AAA','BBB'), drop=c('CCC','DDD')))	
	
	# Check that attempting to specify both keep and drop throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, keep=c('AAA','BBB'), drop=c('CCC','DDD')))	
	
	# Check that attempting to keep a non-existing column (indicated by header name) throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "Conversion", header = TRUE, keep=c('AAA','BBB', 'ZZZ')))	
	
	# Check that attempting to keep a non-existing column (indicated by header name) throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, keep=c('AAA','BBB', 'ZZZ')))
	
	# Check that attempting to keep a column (indicated by index) out of named region bounds throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "Conversion", header = TRUE, keep=c(1,2,5)))	
	
	# Check that attempting to keep a column (indicated by index) out of named region bounds throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, keep=c(1,2,5)))	
	
	# Check that attempting to drop a non-existing column (indicated by header name) throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "Conversion", header = TRUE, drop=c('AAA','BBB', 'ZZZ')))	
	
	# Check that attempting to drop a non-existing column (indicated by header name) throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, drop=c('AAA','BBB', 'ZZZ')))
	
	# Check that attempting to drop a column (indicated by index) out of named region bounds throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "Conversion", header = TRUE, drop=c(1,2,5)))	
	
	# Check that attempting to drop a column (indicated by index) out of named region bounds throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, drop=c(1,2,5)))	
	
	AAA = data.frame(
			A = 1:3,
			B = letters[1:3],
			C = c(TRUE, TRUE, FALSE),
			stringsAsFactors = FALSE
	)
	BBB = data.frame(
			D = 4:6,
			E = letters[4:6],
			F = c(FALSE, TRUE, TRUE),
			stringsAsFactors = FALSE
	)
	
	# Check that keeping columns "NumericColumn" and "BooleanColumn" (= by name) with header=FALSE throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "Test", header = FALSE, keep=c("NumericColumn","BooleanColumn")))
	
	# Check that keeping columns "NumericColumn" and "BooleanColumn" (= by name) with header=FALSE throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "Test", header = FALSE, keep=c("NumericColumn","BooleanColumn")))
	
	# Check that dropping columns "StringColumn" and "DateTimeColumn" (= by name) with header=FALSE throws an exception (*.xls)
	checkException(readNamedRegion(wb.xls, "Test", header = FALSE, drop=c("StringColumn", "DateTimeColumn")))
	
	# Check that dropping columns "StringColumn" and "DateTimeColumn" (= by name) with header=FALSE throws an exception (*.xlsx)
	checkException(readNamedRegion(wb.xlsx, "Test", header = FALSE, drop=c("StringColumn", "DateTimeColumn")))
	
	# Keeping the same columns from multiple named regions (*.xls)
	res <- readNamedRegion(wb.xls, name=c("Test","AAA","BBB"), header = TRUE, keep = c(1,3))
	checkEquals(res, list(Test=checkDf[,c(1,3)], AAA=AAA[,c(1,3)], BBB=BBB[,c(1,3)]))

	# Keeping the same columns from multiple named regions (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name=c("Test","AAA","BBB"), header = TRUE, keep = c(1,3))
	checkEquals(res, list(Test=checkDf[,c(1,3)], AAA=AAA[,c(1,3)], BBB=BBB[,c(1,3)]))
	
	# Testing the correct replication of the keep argument (reading from 3 named regions, while keep has length 2) (*.xls)
	res <- readNamedRegion(wb.xls, name=c("Test","AAA","BBB"), header = TRUE, keep = list(1,3))
	checkEquals(res, list(Test=checkDf[1], AAA=AAA[3], BBB=BBB[1]))
	
	# Testing the correct replication of the keep argument (reading from 3 named regions, while keep has length 2) (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name=c("Test","AAA","BBB"), header = TRUE, keep = list(1,3))
	checkEquals(res, list(Test=checkDf[1], AAA=AAA[3], BBB=BBB[1]))	
	
	# Keeping different columns from multiple named regions (*.xls)
	res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB"), header = TRUE, keep = list(c(1,2),c(2,3),c(1,3)) )
	checkEquals(res, list(Test=checkDf[,c(1,2)], AAA=AAA[,c(2,3)], BBB=BBB[,c(1,3)]))

	# Keeping different columns from multiple named regions (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB"), header = TRUE, keep = list(c(1,2),c(2,3),c(1,3)) )
	checkEquals(res, list(Test=checkDf[,c(1,2)], AAA=AAA[,c(2,3)], BBB=BBB[,c(1,3)]))
		
	#Keeping different columns from multiple named regions (2 keep list elements for 4 named regions) (*.xls)
	res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB", "Test"), header = TRUE, keep = list(c(1,2),c(2,3)))
	checkEquals(res, list(Test=checkDf[,c(1,2)], AAA=AAA[,c(2,3)], BBB=BBB[,c(1,2)], Test=checkDf[,c(2,3)]))

	#Keeping different columns from multiple named regions (2 keep list elements for 4 named regions) (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB", "Test"), header = TRUE, keep = list(c(1,2),c(2,3)))
	checkEquals(res, list(Test=checkDf[,c(1,2)], AAA=AAA[,c(2,3)], BBB=BBB[,c(1,2)], Test=checkDf[,c(2,3)]))
	
	# Dropping the same columns from multiple named regions (*.xls)
	res <- readNamedRegion(wb.xls, name=c("Test", "AAA", "BBB"), header = TRUE, drop = c(1,2))
	checkEquals(res, list(Test=checkDf[,c(3,4)], AAA=data.frame(C=AAA[,3], stringsAsFactors=F), BBB=data.frame(F=BBB[,3], stringsAsFactors=F)))

	# Dropping the same columns from multiple named regions (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name=c("Test", "AAA", "BBB"), header = TRUE, drop = c(1,2))
	checkEquals(res, list(Test=checkDf[,c(3,4)], AAA=data.frame(C=AAA[,3], stringsAsFactors=F), BBB=data.frame(F=BBB[,3], stringsAsFactors=F)))

	# Testing the correct replication of the drop argument (reading from 3 named regions, while drop has length 2) (*.xls)
	res <- readNamedRegion(wb.xls, name=c("Test","AAA","BBB"), header = TRUE, drop = list(1,2))
	checkEquals(res, list(Test=checkDf[,c(2,3,4)], AAA=AAA[,c(1,3)], BBB=BBB[,c(2,3)]))
	
	# Testing the correct replication of the drop argument (reading from 3 named regions, while drop has length 2) (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name=c("Test","AAA","BBB"), header = TRUE, drop = list(1,2))
	checkEquals(res, list(Test=checkDf[,c(2,3,4)], AAA=AAA[,c(1,3)], BBB=BBB[,c(2,3)]))	
	
	# Dropping different columns from multiple named regions (*.xls)
	res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB"), header = TRUE, drop = list(c(1,2),c(2,3),c(1,3)) )
	checkEquals(res, list(Test=checkDf[,c(3,4)], AAA=data.frame(A=AAA[,1], stringsAsFactors=F), BBB=data.frame(E=BBB[,2], stringsAsFactors=F)))
	
	# Dropping different columns from multiple named regions (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB"), header = TRUE, drop = list(c(1,2),c(2,3),c(1,3)) )
	checkEquals(res, list(Test=checkDf[,c(3,4)], AAA=data.frame(A=AAA[,1], stringsAsFactors=F), BBB=data.frame(E=BBB[,2], stringsAsFactors=F)))
	
	#Dropping different columns from multiple named regions (2 drop list elements for 4 named regions) (*.xls)
	res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB", "Test"), header = TRUE, drop = list(c(1,2),c(2,3)))
	checkEquals(res, list(Test=checkDf[,c(3,4)], AAA=data.frame(A=AAA[,1], stringsAsFactors=F), BBB=data.frame(F=BBB[,3], stringsAsFactors=F), Test=checkDf[,c(1,4)]))

	#Dropping different columns from multiple named regions (2 drop list elements for 4 named regions) (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB", "Test"), header = TRUE, drop = list(c(1,2),c(2,3)))
	checkEquals(res, list(Test=checkDf[,c(3,4)], AAA=data.frame(A=AAA[,1], stringsAsFactors=F), BBB=data.frame(F=BBB[,3], stringsAsFactors=F), Test=checkDf[,c(1,4)]))
	
	targetNoForceSubset <- data.frame(
			BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
			DDD = as.POSIXct(c("1984-01-11 12:00:00", NA, NA, NA, NA)),
			stringsAsFactors = FALSE
	)
	
	# Check that conversion performs ok (without forcing conversion, keeping columns BBB and DDD; *.xls)
	res <- readNamedRegion(wb.xls, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S", keep=c('BBB', 'DDD'))
	checkEquals(res, targetNoForceSubset)

	# Check that conversion performs ok (without forcing conversion, keeping columns BBB and DDD; *.xlsx)
	res <- readNamedRegion(wb.xlsx, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S", keep=c('BBB', 'DDD'))
	checkEquals(res, targetNoForceSubset)
	
	# Check that conversion performs ok (without forcing conversion, dropping columns AAA and CCC; *.xls)
	res <- readNamedRegion(wb.xls, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S", drop=c('AAA', 'CCC'))
	checkEquals(res, targetNoForceSubset)
	
	# Check that conversion performs ok (without forcing conversion, dropping columns AAA and CCC; *.xlsx)
	res <- readNamedRegion(wb.xlsx, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S", drop=c('AAA', 'CCC'))
	checkEquals(res, targetNoForceSubset)
	
	# Check that conversion performs ok (without forcing conversion, keeping columns 2 and 4; *.xls)
	res <- readNamedRegion(wb.xls, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S", keep=c(2,4))
	checkEquals(res, targetNoForceSubset)
	
	# Check that conversion performs ok (without forcing conversion, keeping columns 2 and 4; *.xlsx)
	res <- readNamedRegion(wb.xlsx, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S", keep=c(2,4))
	checkEquals(res, targetNoForceSubset)
	
	# Check that conversion performs ok (without forcing conversion, dropping columns 1 and 3; *.xls)
	res <- readNamedRegion(wb.xls, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S", drop=c(1,3))
	checkEquals(res, targetNoForceSubset)
	
	# Check that conversion performs ok (without forcing conversion, dropping columns 1 and 3; *.xlsx)
	res <- readNamedRegion(wb.xlsx, name = "Conversion", header = TRUE,
			colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
			forceConversion = FALSE,
			dateTimeFormat = "%d.%m.%Y %H:%M:%S", drop=c(1,3))
	checkEquals(res, targetNoForceSubset)
  
	# Check that simplification works as expected (*.xls)
  res <- readNamedRegion(wb.xls, name = "Simplify1", header = TRUE, simplify = TRUE)
  checkEquals(res, 1:10)
  res <- readNamedRegion(wb.xls, name = "Simplify2", header = TRUE, simplify = TRUE)
  checkEquals(res, 1:4)
  res <- readNamedRegion(wb.xls, name = "Simplify3", header = TRUE, simplify = TRUE)
  checkEquals(res, c(TRUE, FALSE, FALSE, TRUE))
  res <- readNamedRegion(wb.xls, name = "Simplify4", header = TRUE, simplify = TRUE)
  checkEquals(res, c("one", "two", "three", "four", "five"))
  
	# Check that simplification works as expected (*.xlsx)
	res <- readNamedRegion(wb.xlsx, name = "Simplify1", header = TRUE, simplify = TRUE)
	checkEquals(res, 1:10)
	res <- readNamedRegion(wb.xlsx, name = "Simplify2", header = TRUE, simplify = TRUE)
	checkEquals(res, 1:4)
	res <- readNamedRegion(wb.xlsx, name = "Simplify3", header = TRUE, simplify = TRUE)
	checkEquals(res, c(TRUE, FALSE, FALSE, TRUE))
	res <- readNamedRegion(wb.xlsx, name = "Simplify4", header = TRUE, simplify = TRUE)
	checkEquals(res, c("one", "two", "three", "four", "five"))
  

	# Cached value tests: Create workbook
	wb.xls <- loadWorkbook(rsrc("resources/testCachedValues.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testCachedValues.xlsx"), create = FALSE)

	# "AllLocal" contains no formulae
	ref.xls.uncached <- readNamedRegion(wb.xls, "AllLocal", useCachedValues = FALSE)
	ref.xls.cached <- readNamedRegion(wb.xls, "AllLocal", useCachedValues = TRUE)
	# cached and uncached results should be identical
	checkEquals(ref.xls.uncached, ref.xls.cached)

	ref.xlsx.uncached <- readNamedRegion(wb.xlsx, "AllLocal", useCachedValues = FALSE)
	ref.xlsx.cached <- readNamedRegion(wb.xlsx, "AllLocal", useCachedValues = TRUE)
	checkEquals(ref.xlsx.uncached, ref.xlsx.cached)

	# XLS and XLSX results should be identical
	checkEquals(ref.xlsx.uncached, ref.xls.uncached)

	# the other three named regions reference external worksheets and can't be read
	# with useCachedValues=FALSE
	onErrorCell(wb.xls, XLC$ERROR.STOP)
	checkException(readNamedRegion(wb.xls, "HeaderRemote", useCachedValues = FALSE))
	checkException(readNamedRegion(wb.xls, "BodyRemote", useCachedValues = FALSE))
	checkException(readNamedRegion(wb.xls, "AllRemote", useCachedValues = FALSE))

	onErrorCell(wb.xlsx, XLC$ERROR.STOP)
	checkException(readNamedRegion(wb.xlsx, "HeaderRemote", useCachedValues = FALSE))
	checkException(readNamedRegion(wb.xlsx, "BodyRemote", useCachedValues = FALSE))
	checkException(readNamedRegion(wb.xlsx, "AllRemote", useCachedValues = FALSE))

	res <- readNamedRegion(wb.xls, "HeadersRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
	res <- readNamedRegion(wb.xls, "BodyRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
	res <- readNamedRegion(wb.xls, "BothRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)

	res <- readNamedRegion(wb.xlsx, "HeadersRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
	res <- readNamedRegion(wb.xlsx, "BodyRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
	res <- readNamedRegion(wb.xlsx, "BothRemote", useCachedValues = TRUE)
	checkEquals(ref.xls.uncached, res)
  
	# Check that dimensionality is not dropped when reading in a named region with rownames = x 
	# (see github issue #49)
	expected = data.frame(B = 1:5, row.names = letters[1:5])
	res <- readNamedRegionFromFile(rsrc("resources/testBug49.xlsx"), name = "test", rownames = "A")
	checkEquals(expected, res)
}
