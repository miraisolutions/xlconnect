# Hiding worksheets of an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "hide.xls"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Write a couple of built-in data.frame's into sheets
# with corresponding name
for(obj in c("CO2", "airquality", "swiss")) {
	createSheet(wb, name = obj)
	writeWorksheet(wb, get(obj), sheet = obj)
}

# Hide sheet 'airquality';
# the sheet may be unhidden by a user from within Excel
hideSheet(wb, sheet = "airquality")
# Alternatively: wb$hideSheet(sheet = "airquality")

# Very-hide sheet 'swiss';
# the sheet cannot be unhidden by a user from within Excel
hideSheet(wb, sheet = "swiss", veryHidden = TRUE)
# Alternatively: wb$hideSheet(sheet = "swiss", veryHidden = TRUE)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
# Alternatively: wb$saveWorkbook()

if(interactive() && exists("shell.exec")) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") shell.exec(file.path(getwd(), demoExcelFile))
}
