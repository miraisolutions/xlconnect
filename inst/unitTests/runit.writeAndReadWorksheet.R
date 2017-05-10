#############################################################################
#
# XLConnect
# Copyright (C) 2010-2017 Mirai Solutions GmbH
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
# Tests for checking equality of data.frame's being written to and read
# from Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.writeAndReadWorksheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWriteAndReadWorksheet.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWriteAndReadWorksheet.xlsx"), create = TRUE)
	
	testDataFrame <- function(wb, df, startRow, startCol) {
		worksheet <- deparse(substitute(df))
		createSheet(wb, worksheet)
		writeWorksheet(wb, df, worksheet, startRow = startRow, startCol = startCol)
		res <- readWorksheet(wb, worksheet, startRow = startRow, startCol = startCol,
				endRow = 0, endCol = 0)	
		checkEquals(normalizeDataframe(df), res, check.attributes = FALSE, check.names = TRUE)
	}
	
	if(getOption("FULL.TEST.SUITE")) {
		# built-in dataset mtcars (*.xls)
		testDataFrame(wb.xls, mtcars, 8, 13)
		# built-in dataset mtcars (*.xlsx)
		testDataFrame(wb.xlsx, mtcars, 8, 13)
		
		# built-in dataset airquality (*.xls)
		testDataFrame(wb.xls, airquality, 2, 9)
		# built-in dataset airquality (*.xlsx)
		testDataFrame(wb.xlsx, airquality, 2, 9)
		
		# built-in dataset attenu (*.xls)
		testDataFrame(wb.xls, attenu, 7, 1)
		# built-in dataset attenu (*.xlsx)
		testDataFrame(wb.xlsx, attenu, 7, 1)
		
		# built-in dataset ChickWeight (*.xls)
		testDataFrame(wb.xls, ChickWeight, 8, 8)
		# built-in dataset ChickWeight (*.xlsx)
		testDataFrame(wb.xlsx, ChickWeight, 8, 8)
		
		# built-in dataset CO2 (*.xls)
		testDataFrame(wb.xls, CO2, 100, 12)
		# built-in dataset CO2 (*.xlsx)
		testDataFrame(wb.xlsx, CO2, 100, 12)
		
		# built-in dataset iris (*.xls)
		testDataFrame(wb.xls, iris, 1, 1)
		# built-in dataset iris (*.xlsx)
		testDataFrame(wb.xlsx, iris, 1, 1)
		
		# built-in dataset longley (*.xls)
		testDataFrame(wb.xls, longley, 5, 214)
		# built-in dataset longley (*.xlsx)
		testDataFrame(wb.xlsx, longley, 5, 214)
		
		# built-in dataset morley (*.xls)
		testDataFrame(wb.xls, morley, 1000, 6)
		# built-in dataset morley (*.xlsx)
		testDataFrame(wb.xlsx, morley, 1000, 6)
		
		# built-in dataset swiss (*.xls)
		testDataFrame(wb.xls, swiss, 4, 4)
		# built-in dataset swiss (*.xlsx)
		testDataFrame(wb.xlsx, swiss, 4, 4)
	}
	
	# custom test dataset
	cdf <- data.frame(
			"Column.A" = c(1, 2, 3, NA, 5, 6, 7, 8, NA, 10),
			"Column.B" = c(-4, -3, NA, -1, 0, NA, NA, 3, 4, 5),
			"Column.C" = c("Anna", "???", NA, "", NA, "$!?&%", "(?2@?~?'^*#|)", "{}[]:,;-_<>", "\\sadf\n\nv", "a b c"),
			"Column.D" = c(pi, -pi, NA, sqrt(2), sqrt(0.3), -sqrt(pi), exp(1), log(2), sin(2), -tan(2)),
			"Column.E" = c(TRUE, TRUE, NA, NA, FALSE, FALSE, TRUE, NA, FALSE, TRUE),
			"Column.F" = c("High", "Medium", "Low", "Low", "Low", NA, NA, "Medium", "High", "High"),
			"Column.G" = c("High", "Medium", NA, "Low", "Low", "Medium", NA, "Medium", "High", "High"),
			"Column.H" = rep(c(Sys.Date(), Sys.Date() + 236, NA), length = 10),
			# NOTE: Column.I is automatically converted to POSIXct!!!
			"Column.I" = rep(c(as.POSIXlt(Sys.time()), as.POSIXlt(Sys.time()) + 3523523, NA, as.POSIXlt(Sys.time()) + 838239), length = 10),
			"Column.J" = rep(c(as.POSIXct(Sys.time()), as.POSIXct(Sys.time()) + 436322, NA, as.POSIXct(Sys.time()) - 1295022), length = 10),
			stringsAsFactors = F
	)
	cdf[["Column.F"]] <- factor(cdf[["Column.F"]])
	cdf[["Column.F"]] <- ordered(cdf[["Column.F"]], levels = c("Low", "Medium", "High"))
	
	# (*.xls)
	testDataFrame(wb.xls, cdf, 1, 1)
	# (*.xlsx)
	testDataFrame(wb.xlsx, cdf, 1, 1)
	
	# Check writing of data.frame with row names (*.xls)
	createSheet(wb.xls, "rownames")
	writeWorksheet(wb.xls, mtcars, "rownames", startRow = 9, startCol = 10, header = TRUE, rownames = "Car")
	res <- readWorksheet(wb.xls, "rownames", startRow = 9, startCol = 10)
	checkEquals(XLConnect:::includeRownames(mtcars, "Car"), res, check.attributes = FALSE, check.names = TRUE)
	
	# Check writing of data.frame with row names (*.xlsx)
	createSheet(wb.xlsx, "rownames")
	writeWorksheet(wb.xlsx, mtcars, "rownames", startRow = 9, startCol = 10, header = TRUE, rownames = "Car")
	res <- readWorksheet(wb.xlsx, "rownames", startRow = 9, startCol = 10)
	checkEquals(XLConnect:::includeRownames(mtcars, "Car"), res, check.attributes = FALSE, check.names = TRUE)
	
	# Check writing & reading of data.frame with row names (*.xls)
	createSheet(wb.xls, name = "rownames2")
	writeWorksheet(wb.xls, mtcars, "rownames2", startRow = 31, startCol = 8, header = TRUE, rownames = "Car")
	res <- readWorksheet(wb.xls, "rownames2", rownames = "Car")
	checkEquals(mtcars, res, check.attributes = FALSE, check.names = TRUE)
	checkEquals(attr(mtcars, "row.names"), attr(res, "row.names"))
	
	# Check writing & reading of data.frame with row names (*.xlsx)
	createSheet(wb.xlsx, name = "rownames2")
	writeWorksheet(wb.xlsx, mtcars, "rownames2", startRow = 31, startCol = 8, header = TRUE, rownames = "Car")
	res <- readWorksheet(wb.xlsx, "rownames2", rownames = "Car")
	checkEquals(mtcars, res, check.attributes = FALSE, check.names = TRUE)
	checkEquals(attr(mtcars, "row.names"), attr(res, "row.names"))
}
