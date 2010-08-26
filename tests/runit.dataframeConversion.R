# TODO: Add comment
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
}

