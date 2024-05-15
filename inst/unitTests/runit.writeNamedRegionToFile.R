#############################################################################
#
# XLConnect
# Copyright (C) 2010-2024 Mirai Solutions GmbH
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
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Tests for checking equality of data.frame's being written to and read
# from Excel worksheets
#
# Author: Thomas Themel, Mirai Solutions GmbH
#
#############################################################################

test.writeNamedRegionToFile <- function() {

  if(getOption("FULL.TEST.SUITE")) {
	# Create workbooks
  file.xls <- "testWriteNamedRegionToFileWorkbook.xls"
  file.xlsx <- "testWriteNamedRegionToFileWorkbook.xlsx"

  if(file.exists(file.xls)) {
	file.remove(file.xls)
  }

  if(file.exists(file.xlsx)) {
	file.remove(file.xlsx)
  }

	testDataFrame <- function(file, df) {
		worksheet <- deparse(substitute(df))
	print(paste("Writing dataset ", worksheet, "to file", file))
	name <- paste(worksheet, "Region", sep="")
		writeNamedRegionToFile(file, df, name, formula=paste(worksheet, "A1", sep="!"))
		res <- readNamedRegionFromFile(file, name)
			checkEquals(normalizeDataframe(df), res, check.attributes = FALSE, check.names = TRUE)
	}

	
		# built-in dataset mtcars (*.xls)
		testDataFrame(file.xls, mtcars)
		# built-in dataset mtcars (*.xlsx)
		testDataFrame(file.xlsx, mtcars)

		# built-in dataset airquality (*.xls)
		testDataFrame(file.xls, airquality)
		# built-in dataset airquality (*.xlsx)
		testDataFrame(file.xlsx, airquality)

		# built-in dataset attenu (*.xls)
		testDataFrame(file.xls, attenu)
		# built-in dataset attenu (*.xlsx)
		testDataFrame(file.xlsx, attenu)

		# built-in dataset ChickWeight (*.xls)
		testDataFrame(file.xls, ChickWeight)
		# built-in dataset ChickWeight (*.xlsx)
		testDataFrame(file.xlsx, ChickWeight)

		CO = CO2 # CO2 seems to be an illegal name
		# built-in dataset CO2 (*.xls)
		testDataFrame(file.xls, CO)
		# built-in dataset CO2 (*.xlsx)
		testDataFrame(file.xlsx, CO)

		# built-in dataset iris (*.xls)
		testDataFrame(file.xls, iris)
		# built-in dataset iris (*.xlsx)
		testDataFrame(file.xlsx, iris)

		# built-in dataset longley (*.xls)
		testDataFrame(file.xls, longley)
		# built-in dataset longley (*.xlsx)
		testDataFrame(file.xlsx, longley)

		# built-in dataset morley (*.xls)
		testDataFrame(file.xls, morley)
		# built-in dataset morley (*.xlsx)
		testDataFrame(file.xlsx, morley)

		# built-in dataset swiss (*.xls)
		testDataFrame(file.xls, swiss)
		# built-in dataset swiss (*.xlsx)
		testDataFrame(file.xlsx, swiss)

	# custom test dataset
	cdf <- data.frame(
			"Column.A" = c(1, 2, 3, NA, 5, 6, 7, 8, NA, 10),
			"Column.B" = c(-4, -3, NA, -1, 0, NA, NA, 3, 4, 5),
			"Column.C" = c("Anna", "???", NA, "", NA, "$!?&%", "(?2@?~?'^*#|)", "{}[]:,;-_<>", "\\sadf\n\nv", "a b c"),
			"Column.D" = c(pi, -pi, NA, sqrt(2), sqrt(0.3), -sqrt(pi), exp(1), log(2), sin(2), -tan(2)),
			"Column.E" = c(TRUE, TRUE, NA, NA, FALSE, FALSE, TRUE, NA, FALSE, TRUE),
			"Column.F" = c("High", "Medium", "Low", "Low", "Low", NA, NA, "Medium", "High", "High"),
			"Column.G" = c("High", "Medium", NA, "Low", "Low", "Medium", NA, "Medium", "High", "High"),
			"Column.H" = rep(c(as.Date("2021-10-30") , as.Date("2021-03-28"), NA), length = 10),
			# NOTE: Column.I is automatically converted to POSIXct!!!
			"Column.I" = rep(c(as.POSIXlt("2021-10-31 03:00:00"), as.POSIXlt(1582963631, origin="1970-01-01"), NA, as.POSIXlt("2001-12-31 23:59:59")), length = 10),
			# NOTE: 1582963631 with origin="1970-01-01" corresponds to 2020 Feb 29
			"Column.J" = rep(c(as.POSIXct("2021-10-31 03:00:00"), as.POSIXct(1582963631, origin="1970-01-01"), NA, as.POSIXct("2001-12-31 23:59:59")), length = 10),
			stringsAsFactors = F
	)
	cdf[["Column.F"]] <- factor(cdf[["Column.F"]])
	cdf[["Column.F"]] <- ordered(cdf[["Column.F"]], levels = c("Low", "Medium", "High"))

	# (*.xls)
	testDataFrame(file.xls, cdf)
	# (*.xlsx)
	testDataFrame(file.xlsx, cdf)

	file2.xls <- "wnrtf1.xls"
	file2.xlsx <- "wnrtf1.xlsx"
	if(file.exists(file2.xls)) file.remove(file2.xls)
	if(file.exists(file2.xlsx)) file.remove(file2.xlsx)
	
	# Check that writing a data.frame to a named region with a formula that contains a white space in
	# the sheet name does not cause any grief (*.xls)
	checkNoException(writeNamedRegionToFile(file2.xls, data = mtcars, name = "mtcars",
		formula = "'My Cars'!$A$1", header = TRUE))
	checkTrue(file.exists(file2.xls))

	# Check that writing a data.frame to a named region with a formula that contains a white space in
	# the sheet name does not cause any grief (*.xlsx)
	checkNoException(writeNamedRegionToFile(file2.xlsx, data = mtcars, name = "mtcars",
					formula = "'My Cars'!$A$1", header = TRUE))
	checkTrue(file.exists(file2.xlsx))

  # test clearNamedRegions
  testClearNamedRegions <- function(file, df) {
	df.short <- df[1,]

	# overwrite named region with shorter version
	writeNamedRegionToFile(file, data=df.short, name="cdfRegion")
	# default behaviour: not cleared, only named region is shortened
	checkEquals(nrow(readNamedRegionFromFile(file, name="cdfRegion")), 1)
	checkEquals(nrow(readWorksheetFromFile(file, sheet="cdf")), nrow(df))

	# rewrite longer version
	writeNamedRegionToFile(file, data=df, name="cdfRegion")
	# overwrite name with shorter version & clearing
	writeNamedRegionToFile(file, data=df.short, name="cdfRegion", clearNamedRegions=TRUE)
	# should be cleared
	checkEquals(nrow(readNamedRegionFromFile(file, name="cdfRegion")), 1)
	checkEquals(nrow(readWorksheetFromFile(file, sheet="cdf")), 1)
  }

  testClearNamedRegions(file.xls, cdf)
  testClearNamedRegions(file.xlsx, cdf)

  testClearNamedRegionsScoped <- function(file, df) {
	scope <- c('scope1', 'scope2')
	clearParam <- c(TRUE, FALSE)
	df.short <- df[1,]

	wb <- loadWorkbook(file, create = TRUE)
	createSheet(wb, scope)
	saveWorkbook(wb, file)

	writeNamedRegionToFile(file, data=df, name="cdfRegionScoped", formula=paste(scope, "A1", sep="!"), worksheetScope = scope) # should write in cell A1 in each sheet
	# overwrite named region with shorter version
	writeNamedRegionToFile(file, data=df.short, name="cdfRegionScoped", worksheetScope = scope)
	# default behaviour: not cleared, only named region is shortened
	checkEquals(nrow(readNamedRegionFromFile(file, name="cdfRegionScoped", worksheetScope = scope)[[1]]), 1)
	checkEquals(nrow(readNamedRegionFromFile(file, name="cdfRegionScoped", worksheetScope = scope)[[2]]), 1)
	checkEquals(nrow(readWorksheetFromFile(file, sheet=scope)[[1]]), nrow(df))
	checkEquals(nrow(readWorksheetFromFile(file, sheet=scope)[[2]]), nrow(df))

	# rewrite longer version 
	writeNamedRegionToFile(file, data=df, name="cdfRegionScoped", worksheetScope = scope)
	# overwrite name with shorter version & clearing
	writeNamedRegionToFile(file, data=df.short, name="cdfRegionScoped", clearNamedRegions=clearParam, worksheetScope = scope)
	# should be cleared
	checkEquals(nrow(readNamedRegionFromFile(file, name="cdfRegionScoped", worksheetScope = scope)[[1]]), 1)
	checkEquals(nrow(readNamedRegionFromFile(file, name="cdfRegionScoped", worksheetScope = scope)[[2]]), 1)
	checkEquals(nrow(readWorksheetFromFile(file, sheet=scope)[[1]]), 1)
	checkEquals(nrow(readWorksheetFromFile(file, sheet=scope)[[2]]), nrow(df))
  }

  scopedfile.xls <- "testWriteNamedRegionToFileWorkbookScoped.xls"
  scopedfile.xlsx <- "testWriteNamedRegionToFileWorkbookScoped.xlsx"
  
  testClearNamedRegionsScoped(scopedfile.xls, cdf)
  testClearNamedRegionsScoped(scopedfile.xlsx, cdf)
  }
}
