# Writing a named region to an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "mtcars.xlsx"

# Open workbook (create if not existing)
wb <- openWorkbook(demoExcelFile, create = TRUE)

# Create a named region called 'mtcars' on a sheet called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$A$1")
# Alternatively: wb$createName(name = "mtcars", formula = "mtcars!$A$1")

# Write built-in data set 'mtcars' to the above defined named region
writeNamedRegion(wb, mtcars, name = "mtcars")
# Alternatively: wb$writeNamedRegion(mtcars, name = "mtcars")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
# Alternatively: wb$saveWorkbook()
