# Creating a summary of an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(XLConnect)

# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Show workbook summary
summary(wb)
