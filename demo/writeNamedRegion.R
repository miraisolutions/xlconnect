# Writing a named region to an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "mtcars.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Create a named region called 'mtcars' on a sheet called 'mtcars'
# (the call to 'createName' automatically creates the sheet
# referenced in the formula if it does not exist)
createName(wb, name = "mtcars", formula = "mtcars!$A$1")
# Alternatively: wb$createName(name = "mtcars", formula = "mtcars!$A$1")

# Write built-in data set 'mtcars' to the above defined named region
writeNamedRegion(wb, mtcars, name = "mtcars")
# Alternatively: wb$writeNamedRegion(mtcars, name = "mtcars")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
# Alternatively: wb$saveWorkbook()

if(interactive() && exists("shell.exec")) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") shell.exec(file.path(getwd(), demoExcelFile))
}
