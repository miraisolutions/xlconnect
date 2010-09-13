# Writing a worksheet to an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "CO2.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Create a worksheet called 'CO2'
createSheet(wb, name = "CO2")
# Alternatively: wb$createSheet(name = "CO2")

# Write built-in data set 'CO2' to the worksheet created above;
# offset from the top left corner and with default header = TRUE
writeWorksheet(wb, CO2, sheet = "CO2", startRow = 4, startCol = 2)
# Alternatively: wb$writeNamedRegion(CO2, sheet = "CO2", startRow = 4, startCol = 2)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
# Alternatively: wb$saveWorkbook()

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") shell.exec(file.path(getwd(), demoExcelFile))
}
