# Tests whether data.frame's pushed and pulled to/from Java
# are consistent (with respect to some defined differences;
# see 'normalizeDataframe')
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.dataframeConversion <- function() {
	
	# built-in dataset mtcars
	res <- dataframeFromJava(dataframeToJava(mtcars))
	checkEquals(normalizeDataframe(mtcars), res)
	
	# built-in dataset airquality
	res <- dataframeFromJava(dataframeToJava(airquality))
	checkEquals(normalizeDataframe(airquality), res)
	
	# built-in dataset attenu
	res <- dataframeFromJava(dataframeToJava(attenu))
	checkEquals(normalizeDataframe(attenu), res)
	
	# built-in dataset ChickWeight
	res <- dataframeFromJava(dataframeToJava(ChickWeight))
	checkEquals(normalizeDataframe(ChickWeight), res)
	
	# built-in dataset CO2
	res <- dataframeFromJava(dataframeToJava(CO2))
	checkEquals(normalizeDataframe(CO2), res)
	
	# built-in dataset iris
	res <- dataframeFromJava(dataframeToJava(iris))
	checkEquals(normalizeDataframe(iris), res)
	
	# built-in dataset longley
	res <- dataframeFromJava(dataframeToJava(longley))
	checkEquals(normalizeDataframe(longley), res)
	
	# built-in dataset morley
	res <- dataframeFromJava(dataframeToJava(morley))
	checkEquals(normalizeDataframe(morley), res)
	
	# built-in dataset swiss
	res <- dataframeFromJava(dataframeToJava(swiss))
	checkEquals(normalizeDataframe(swiss), res)
	
	# custom test dataset
	cdf <- data.frame(
			"Column.A" = c(1, 2, 3, NA, 5, 6, 7, 8, NA, 10),
			"Column.B" = c(-4, -3, NA, -1, 0, NA, NA, 3, 4, 5),
			"Column.C" = c("Anna", "äöü", 
					NA, "", NA, "$!?&%", "(£2@§~°'^*#|)", "{}[]:,;-_<>", "\\sadf\n\nv", "a b c"),
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
	
	res <- dataframeFromJava(dataframeToJava(cdf))
	checkEquals(normalizeDataframe(cdf), res)
}

