# Reading a named region from an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(XLConnect)

# mtcars xlsx file from demo subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Read named region 'mtcars' (with default header = TRUE)
data <- readNamedRegion(wb, name = "mtcars")
# Alternatively: wb$readNamedRegion(name = "mtcars")

# Print resulting data.frame
print(data)
