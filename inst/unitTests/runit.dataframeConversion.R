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
# Tests whether data.frame's pushed and pulled to/from Java
# are consistent (with respect to some defined differences;
# see 'normalizeDataframe')
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.dataframeConversion <- function() {
	
	testDataFrame <- function(df) {
		res <- XLConnect:::dataframeFromJava(XLConnect:::dataframeToJava(df), check.names = TRUE)
		checkEquals(normalizeDataframe(df), res, check.attributes = FALSE, check.names = TRUE)
	}
	
	if(getOption("FULL.TEST.SUITE")) {
		# built-in dataset mtcars
		testDataFrame(mtcars)
		
		# built-in dataset airquality
		testDataFrame(airquality)
		
		# built-in dataset attenu
		testDataFrame(attenu)
		
		# built-in dataset ChickWeight
		testDataFrame(ChickWeight)
		
		# built-in dataset CO2
		testDataFrame(CO2)
		
		# built-in dataset iris
		testDataFrame(iris)
		
		# built-in dataset longley
		testDataFrame(longley)
		
		# built-in dataset morley
		testDataFrame(morley)
		
		# built-in dataset swiss
		testDataFrame(swiss)
	}
	
	# custom test dataset
	cdf <- data.frame(
			"Column.A" = c(1, 2, 3, NA, 5, Inf, 7, 8, NA, 10),
			"Column.B" = c(-4, -3, NA, -Inf, 0, NA, NA, 3, 4, 5),
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
	
	testDataFrame(cdf)
	
	# Check that when being supplied with an object that is not coercable
	# into a data.frame, an appropriate exception is thrown
	checkException(XLConnect:::dataframeToJava(search))
	
	# Check that exceptions are thrown when calling dataframeFromJava
	# with inappropriate objects
	checkException(XLConnect:::dataframeFromJava(NULL))
	checkException(XLConnect:::dataframeFromJava(NA))
	checkException(XLConnect:::dataframeFromJava(9))
	checkException(XLConnect:::dataframeFromJava(search))
}

