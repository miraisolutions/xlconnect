### R code from vignette source 'vignettes/XLConnect.Rnw'
### Encoding: UTF-8

###################################################
### code chunk number 1: setup
###################################################
if (!file.exists('figs')) dir.create('figs')


###################################################
### code chunk number 2: load
###################################################
require(XLConnect)


###################################################
### code chunk number 3: removeFile1
###################################################
if (file.exists("XLConnectExample1.xlsx"))
  file.remove("XLConnectExample1.xlsx")


###################################################
### code chunk number 4: simpleEx1
###################################################
require(XLConnect)
wb <- loadWorkbook("XLConnectExample1.xlsx", create = TRUE)
createSheet(wb, name = "chickSheet")
writeWorksheet(wb, ChickWeight, sheet = "chickSheet", startRow = 3, startCol = 4)
saveWorkbook(wb)


###################################################
### code chunk number 5: removeFile2
###################################################
if (file.exists("XLConnectExample2.xlsx"))
  file.remove("XLConnectExample2.xlsx")


###################################################
### code chunk number 6: simpleEx2
###################################################
require(XLConnect)
writeWorksheetToFile("XLConnectExample2.xlsx", data = ChickWeight,
                     sheet = "chickSheet", startRow = 3, startCol = 4)


###################################################
### code chunk number 7: removeFile3
###################################################
if (file.exists("XLConnectExample3.xlsx"))
  file.remove("XLConnectExample3.xlsx")


###################################################
### code chunk number 8: simpleEx3
###################################################
require(XLConnect)
wb <- loadWorkbook("XLConnectExample3.xlsx", create = TRUE)
createSheet(wb, name = "womenData")
createName(wb, name = "womenName", formula = "womenData!$C$5", overwrite = TRUE)
writeNamedRegion(wb, women, name = "womenName")
saveWorkbook(wb)


###################################################
### code chunk number 9: removeFile4
###################################################
if (file.exists("XLConnectExample4.xlsx"))
  file.remove("XLConnectExample4.xlsx")


###################################################
### code chunk number 10: simpleEx4
###################################################
require(XLConnect)
writeNamedRegionToFile("XLConnectExample4.xlsx", women, name = "womenName",
                       formula = "womenData!$C$5")


###################################################
### code chunk number 11: latexEx1
###################################################
require(XLConnect)
wb <- loadWorkbook("XLConnectExample1.xlsx", create = TRUE)
data <- readWorksheet(wb, sheet = "chickSheet", startRow = 0, endRow = 10,
                      startCol = 0, endCol = 0)
data


###################################################
### code chunk number 12: simpleEx5
###################################################
require(XLConnect)
data <- readWorksheetFromFile("XLConnectExample1.xlsx", sheet = "chickSheet",
                              startRow = 0, endRow = 10, startCol = 0, endCol = 0)


###################################################
### code chunk number 13: latexEx2
###################################################
require(XLConnect)
wb <- loadWorkbook("XLConnectExample3.xlsx", create = TRUE)
data <- readNamedRegion(wb, name = "womenName")
data


###################################################
### code chunk number 14: simpleEx6
###################################################
require(XLConnect)
data <- readNamedRegionFromFile("XLConnectExample3.xlsx", "womenName")


###################################################
### code chunk number 15: AdvancedExampleP1
###################################################
require(XLConnect)
require(zoo)
require(ggplot2) # >= 0.9.3


###################################################
### code chunk number 16: AdvancedExampleP2
###################################################
# Read in (historical) currency data
# For this example take sample data set 'swissfranc' from XLConnect
curr <- XLConnect::swissfranc
curr <- curr[order(curr$Date),]


###################################################
### code chunk number 17: removeFile2
###################################################
if (file.exists("swiss_franc.xlsx"))
  file.remove("swiss_franc.xlsx")


###################################################
### code chunk number 18: AdvancedExampleP3
###################################################

# Workbook filename
wbFilename <- "swiss_franc.xlsx"

# Create a new workbook
wb <- loadWorkbook(wbFilename, create = TRUE)

# Create a new sheet named 'Swiss_Franc'
sheet <- "Swiss_Franc"
createSheet(wb, name = sheet)

# Create a new Excel name referring to the top left corner
# of the sheet 'Swiss_Franc' - this name is going to hold
# our currency data
dataName <- "currency"
nameLocation <- paste(sheet, "$A$1", sep = "!")
createName(wb, name = dataName, formula = nameLocation)

# Instruct XLConnect to only apply a data format for a cell
# but not to apply any other cell styling
setStyleAction(wb, XLC$"STYLE_ACTION.DATA_FORMAT_ONLY")

# Set the default format for numeric data to display
# four digits after the decimal point
setDataFormatForType(wb, type = XLC$"DATA_TYPE.NUMERIC", format = "0.0000")

# Write the currency data to the named region created above
# Note: the named region will be automatically redefined to encompass all
# written data
writeNamedRegion(wb, data = curr, name = dataName, header = TRUE)

# Save the workbook (this actually writes the file to disk)
saveWorkbook(wb)


###################################################
### code chunk number 19: AdvancedExampleP4
###################################################

# Load the workbook created above
wb <- loadWorkbook(wbFilename)

# Create a cell style for the header row
csHeader <- createCellStyle(wb, name = "header")
setFillPattern(csHeader, fill = XLC$FILL.SOLID_FOREGROUND)
setFillForegroundColor(csHeader, color = XLC$COLOR.GREY_25_PERCENT)

# Create a date cell style with a custom format for the Date column
csDate <- createCellStyle(wb, name = "date")
setDataFormat(csDate, format = "yyyy-mm-dd")

# Create a highlighting cell style
csHlight <- createCellStyle(wb, name = "highlight")
setFillPattern(csHlight, fill = XLC$FILL.SOLID_FOREGROUND)
setFillForegroundColor(csHlight, color = XLC$COLOR.CORNFLOWER_BLUE)

# Apply header cell style to the header row
setCellStyle(wb, sheet = sheet, row = 1,  
             col = seq(length.out = ncol(curr)),
             cellstyle = csHeader)

# Index for all rows except header row
allRows <- seq(length = nrow(curr)) + 1

# Apply date cell style to the Date column
setCellStyle(wb, sheet = sheet, row = allRows, col = 1, 
             cellstyle = csDate)

# Set column width such that the full date column is visible
setColumnWidth(wb, sheet = sheet, column = 1, width = 2800)

# Check if there was a change of more than 2% compared 
# to the previous day (per currency)
idx <- rollapply(curr[, -1], width = 2,
                 FUN = function(x) abs(x[2] / x[1] - 1),
                 by.column = TRUE) > 0.02

idx <- rbind(rep(FALSE, ncol(idx)), idx)
widx <- lapply(as.data.frame(idx), which)
# Apply highlighting cell style
for (i in seq(along = widx)) {
  if (length(widx[[i]]) > 0) {
    setCellStyle(wb, sheet = sheet, row = widx[[i]] + 1, col = i + 1,
                 cellstyle = csHlight)
  }

  # Note:
  # +1 for row since there is a header row
  # +1 for column since the first column is the time column
}

saveWorkbook(wb)


###################################################
### code chunk number 20: AdvancedExampleP5
###################################################
wb <- loadWorkbook(wbFilename)

# Stack currencies into a currency variable (for use with ggplot2 below)
currencies <- names(curr)[-1]
gcurr <- reshape(curr, varying = currencies, direction = "long", v.names = "Value",
                 times = currencies, timevar = "Currency")

# Create a png graph showing the currencies in the context 
# of the Swiss Franc
png(filename = "swiss_franc.png", width = 800, height = 600)
p <- ggplot(gcurr, aes(Date, Value, colour = Currency)) +
  geom_line() + stat_smooth(method = "loess") +
  scale_y_continuous("Exchange Rate CHF/CUR") +
  labs(title = paste0("CHF vs ", paste(currencies, collapse = ", ")), 
       x = "") +
  theme(axis.title.y = element_text(size = 10, angle = 90, vjust = 0.3))
print(p)
dev.off()

# Define where the image should be placed via a named region;
# let's put the image two columns left to the data starting 
# in the 5th row
createName(wb, name = "graph",
formula <- paste(sheet, idx2cref(c(5, ncol(curr) + 2)), sep = "!"))
# Note: idx2cref converts indices (row, col) to Excel cell references

# Put the image created above at the corresponding location
addImage(wb, filename = "swiss_franc.png", name = "graph",
         originalSize = TRUE)

saveWorkbook(wb)


###################################################
### code chunk number 21: XLConnect.Rnw:756-757
###################################################
p


