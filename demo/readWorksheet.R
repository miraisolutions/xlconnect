# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(XLConnect)

# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

## CASE 1: 
## Data starting in top left corner; no other data
## contained on same worksheet

# Read worksheet 'mtcars' (providing no specific area bounds;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars")
# Alternatively: wb$readWorksheet(sheet = "mtcars")

# Print resulting data.frame
print(data)

## CASE 2: 
## Data offset from top left corner; no other data
## contained on same worksheet

# Read worksheet 'mtcars2' (providing no specific area bounds;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars2")
# Alternatively: wb$readWorksheet(sheet = "mtcars2")

# Print resulting data.frame
print(data)

## CASE 3: 
## Data offset from top left corner; also, other data contained
## on same worksheet

# Read worksheet 'mtcars3' (providing specific area bounds;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars3", startRow = 10, startCol = 6,
		endRow = 42, endCol = 16)
# Alternatively: wb$readWorksheet(sheet = "mtcars3", startRow = 10, startCol = 6,
#					endRow = 42, endCol = 16)

# Print resulting data.frame
print(data)
