# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


remove(list = objects())
excelFile <- "C:/temp/test.xls"
file.remove(excelFile)

library(XLConnect)
wb <- openWorkbook(excelFile, create = TRUE)
df.in <- data.frame(
		A = c("A", "B", NA, "D"), 
		B = c(1, NA, 3, 4),
		C = c(NA, TRUE, FALSE, TRUE))
writeNamedRegion(wb, df.in, "Test", "Test!$A$1", TRUE)
saveWorkbook(wb)

df.out <- readNamedRegion(wb, "Test", TRUE)


readWorksheet(wb, worksheet = 1, header = TRUE)

existsName(wb, "Test")
existsName(wb, "asdf")

existsSheet(wb, "Test")
existsSheet(wb, "xxx")

writeWorksheet(wb, data = mtcars, worksheet = "mtcars", create = TRUE)
