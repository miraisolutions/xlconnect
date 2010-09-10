# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.writeAndReadNamedRegion <- function() {
	
	# Create workbooks
	wb.xls <- openWorkbook("resources/testWriteAndReadNamedRegion.xls", create = TRUE)
	wb.xlsx <- openWorkbook("resources/testWriteAndReadNamedRegion.xlsx", create = TRUE)
	
	testDataFrame <- function(wb, df, lref) {
		namedRegion <- deparse(substitute(df))
		createName(wb, name = namedRegion, formula = paste(namedRegion, lref, sep = "!"))
		writeNamedRegion(wb, df, name = namedRegion, header = TRUE)
		res <- readNamedRegion(wb, namedRegion)
		checkEquals(normalizeDataframe(df), res)
	}
	
	# built-in dataset mtcars (*.xls)
	testDataFrame(wb.xls, mtcars, "$C$8")
	# built-in dataset mtcars (*.xlsx)
	testDataFrame(wb.xlsx, mtcars, "$C$8")
	
	# built-in dataset airquality (*.xls)
	testDataFrame(wb.xls, airquality, "$F$13")
	# built-in dataset airquality (*.xlsx)
	testDataFrame(wb.xlsx, airquality, "$F$13")
	
	# built-in dataset attenu (*.xls)
	testDataFrame(wb.xls, attenu, "$A$8")
	# built-in dataset attenu (*.xlsx)
	testDataFrame(wb.xlsx, attenu, "$A$8")
	
	# built-in dataset ChickWeight (*.xls)
	testDataFrame(wb.xls, ChickWeight, "$BQ$7")
	# built-in dataset ChickWeight (*.xlsx)
	testDataFrame(wb.xlsx, ChickWeight, "$BQ$7")
	
	# built-in dataset CO2 (*.xls)
	testDataFrame(wb.xls, CO2, "$L$1")
	# built-in dataset CO2 (*.xlsx)
	testDataFrame(wb.xlsx, CO2, "$L$1")
	
	# built-in dataset iris (*.xls)
	testDataFrame(wb.xls, iris, "$BB$5")
	# built-in dataset iris (*.xlsx)
	testDataFrame(wb.xlsx, iris, "$BB$5")
	
	# built-in dataset longley (*.xls)
	testDataFrame(wb.xls, longley, "$AD$8")
	# built-in dataset longley (*.xlsx)
	testDataFrame(wb.xlsx, longley, "$AD$8")
	
	# built-in dataset morley (*.xls)
	testDataFrame(wb.xls, morley, "$K$4")
	# built-in dataset morley (*.xlsx)
	testDataFrame(wb.xlsx, morley, "$K$4")
	
	# built-in dataset swiss (*.xls)
	testDataFrame(wb.xls, swiss, "$M$2")
	# built-in dataset swiss (*.xlsx)
	testDataFrame(wb.xlsx, swiss, "$M$2")
	
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
	
	# (*.xls)
	testDataFrame(wb.xls, cdf, "$X$100")
	# (*.xlsx)
	testDataFrame(wb.xlsx, cdf, "$X$100")
	
}
