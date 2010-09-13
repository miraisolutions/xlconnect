# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################


remove(list = objects())
excelFile <- "C:/temp/mtcars.xlsx"
# excelFile <- "C:/Users/mstuder/Documents/testWorkbookRemoveSheet.xls"
file.remove(excelFile)

library(XLConnect)
wb <- openWorkbook(excelFile, create = TRUE)
# wb <- openWorkbook(excelFile, create = FALSE)
df.in <- data.frame(
		A = c("A", "B", NA, "D"), 
		B = c(1, NA, 3, 4),
		C = c(NA, TRUE, FALSE, TRUE))
createName(wb, name = "Test", formula = "Test!$A$1", overwrite = TRUE)
writeNamedRegion(wb, df.in, name = "Test", header = TRUE)
saveWorkbook(wb)
