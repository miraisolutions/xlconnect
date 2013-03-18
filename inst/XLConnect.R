### R code from vignette source 'XLConnect.Rnw'
### Encoding: UTF-8

###################################################
### code chunk number 1: setup
###################################################

	if( !file.exists( 'figs' ) ) dir.create( 'figs' )
	# require(tikzDevice)
	


###################################################
### code chunk number 2: setup
###################################################
	require(XLConnect)


###################################################
### code chunk number 3: removeFile1
###################################################
	if(file.exists("XLConnectExample1.xlsx"))
		file.remove("XLConnectExample1.xlsx")


###################################################
### code chunk number 4: simpleEx
###################################################
require(XLConnect)
wb <- loadWorkbook("XLConnectExample1.xlsx", create = TRUE)
createSheet(wb, name = "chickSheet")
writeWorksheet(wb, ChickWeight, sheet = "chickSheet", startRow = 3, startCol = 4)
saveWorkbook(wb)


###################################################
### code chunk number 5: removeFile1
###################################################
	if(file.exists("XLConnectExample2.xlsx"))
		file.remove("XLConnectExample2.xlsx")


###################################################
### code chunk number 6: simpleEx
###################################################
require(XLConnect)
writeWorksheetToFile("XLConnectExample2.xlsx", data = ChickWeight, 
sheet = "chickSheet", startRow = 3, startCol = 4)


###################################################
### code chunk number 7: removeFile3
###################################################
	if(file.exists("XLConnectExample3.xlsx"))
		file.remove("XLConnectExample3.xlsx")


###################################################
### code chunk number 8: simpleEx
###################################################
require(XLConnect)
wb = loadWorkbook("XLConnectExample3.xlsx", create = TRUE)
createSheet(wb, name = "womenData")
createName(wb, name = "womenName", formula = "womenData!$C$5", overwrite = TRUE)
writeNamedRegion(wb, women, name = "womenName")
saveWorkbook(wb)


###################################################
### code chunk number 9: removeFile4
###################################################
	if(file.exists("XLConnectExample4.xlsx"))
		file.remove("XLConnectExample4.xlsx")


###################################################
### code chunk number 10: simpleEx
###################################################
require(XLConnect)
writeNamedRegionToFile("XLConnectExample4.xlsx", women, 
		name = "womenName", formula = "womenData!$C$5")


###################################################
### code chunk number 11: latexEx
###################################################
require(XLConnect)
wb = loadWorkbook("XLConnectExample1.xlsx", create = TRUE)
data = readWorksheet(wb, sheet = "chickSheet", startRow = 0, endRow = 10, 
		startCol = 0, endCol = 0)
data


###################################################
### code chunk number 12: simpleEx
###################################################
require(XLConnect)
data = readWorksheetFromFile("XLConnectExample1.xlsx", 
		sheet = "chickSheet", startRow = 0, endRow = 10, 
		startCol = 0, endCol = 0)


###################################################
### code chunk number 13: latexEx
###################################################
require(XLConnect)
wb = loadWorkbook("XLConnectExample3.xlsx", create = TRUE)
data = readNamedRegion(wb, name = "womenName")
data


###################################################
### code chunk number 14: simpleEx
###################################################
require(XLConnect)
data = readNamedRegionFromFile("XLConnectExample3.xlsx", "womenName")


###################################################
### code chunk number 15: AdvancedExampleP1
###################################################
	require(XLConnect)
	require(fImport)
	require(forecast)
	require(zoo)
	require(ggplot2) # >= 0.9.3
  require(scales)


###################################################
### code chunk number 16: AdvancedExampleP2
###################################################
# Currencies we're interested in compared to CHF
currencies = c("EUR", "USD", "GBP", "JPY")

# Fetch currency exchange rates (currency to CHF) 
# from OANDA (last 366 days)
curr = do.call("cbind", args = lapply(currencies,
               function(cur) oandaSeries(paste(cur, "CHF", sep = "/"))))
# Make a copy for later use
curr.orig = curr
# Scale currencies to exchange rate on first day in the series (baseline)
curr = curr * matrix(1/curr[1,], nrow = nrow(curr),
                     ncol = ncol(curr), byrow = TRUE) - 1
# Some data transformations to bring the data into a simple data.frame
curr = transform(curr, Time = time(curr)@Data)
names(curr) = c(currencies, "Time")
# Cyclic shift to bring the Time column to the front
curr = curr[(seq(along = curr) - 2) %% ncol(curr) + 1]

# Number of days to predict
predictDays = 20
# For each currency ...
currFit = sapply(curr[, -1], function(cur) {
  as.numeric(forecast(cur, h = predictDays)$mean)
})
# Add Time column to predictions
currFit = cbind(
  Time = seq(from = curr[nrow(curr), "Time"],
             length.out = predictDays + 1, by = "days")[-1],
  as.data.frame(currFit))

# Bind actual data with predictions
curr = rbind(curr, currFit)


###################################################
### code chunk number 17: removeFile2
###################################################
	if(file.exists("swiss_franc.xlsx"))
		file.remove("swiss_franc.xlsx")


###################################################
### code chunk number 18: AdvancedExampleP3
###################################################

# Workbook filename
wbFilename = "swiss_franc.xlsx"
# Create a new workbook
wb = loadWorkbook(wbFilename, create = TRUE)

# Create a new sheet named 'Swiss_Franc'
sheet = "Swiss_Franc"
createSheet(wb, name = sheet)

# Create a new Excel name referring to the top left corner
# of the sheet 'Swiss_Franc' - this name is going to hold
# our currency data
dataName = "currency"
nameLocation = paste(sheet, "$A$1", sep = "!")
createName(wb, name = dataName, formula = nameLocation)

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
wb = loadWorkbook(wbFilename)

# Create a date cell style with a custom format for the Time column
# (only show year, month and day without any time fields)
csDate = createCellStyle(wb, name = "date")
setDataFormat(csDate, format = "yyyy-mm-dd")
# Create a time/date cell style for the prediction records
csPrediction = createCellStyle(wb, name = "prediction")
setDataFormat(csPrediction, format = "yyyy-mm-dd")
setFillPattern(csPrediction, fill = XLC$FILL.SOLID_FOREGROUND)
setFillForegroundColor(csPrediction, color = XLC$COLOR.GREY_25_PERCENT)
# Create a percentage cell style
# Number format: 2 digits after decimal point
csPercentage = createCellStyle(wb, name = "currency")
setDataFormat(csPercentage, format = "0.00%")
# Create a highlighting cell style
csHlight = createCellStyle(wb, name = "highlight")
setFillPattern(csHlight, fill = XLC$FILL.SOLID_FOREGROUND)
setFillForegroundColor(csHlight, color = XLC$COLOR.CORNFLOWER_BLUE)
setDataFormat(csHlight, format = "0.00%")

# Index for all rows except header row
allRows = seq(length = nrow(curr)) + 1

# Apply date cell style to the Time column
setCellStyle(wb, sheet = sheet, row = allRows, col = 1, 
		cellstyle = csDate)
# Set column width such that the full date column is visible
setColumnWidth(wb, sheet = sheet, column = 1, width = 2800)
# Apply prediction cell style
setCellStyle(wb, sheet = sheet, row = tail(allRows, n = predictDays), 
		col = 1, cellstyle = csPrediction)
# Apply number format to the currency columns
currencyColumns = seq(along = currencies) + 1
for(col in currencyColumns) {
	setCellStyle(wb, sheet = sheet, row = allRows, col = col,
	cellstyle = csPercentage)
}

# Check if there was a change of more than 2% compared 
# to the previous day (per currency)
idx = rollapply(curr.orig, width = 2, 
		FUN = function(x) abs(x[2] / x[1] - 1),
by.column = TRUE) > 0.02
idx = rbind(rep(FALSE, ncol(idx)), idx)
widx = lapply(as.data.frame(idx), which)
# Apply highlighting cell style
for(i in seq(along = currencies)) {
	if(length(widx[[i]]) > 0) {
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
wb = loadWorkbook(wbFilename)

# Stack currencies into a currency variable (for use with ggplot2 below)
gcurr = reshape(curr, varying = currencies, direction = "long",
v.names = "Value", times = currencies, timevar = "Currency")

# Also add a discriminator column to differentiate between actual and
# prediction values
gcurr[["Type"]] = ifelse(gcurr$Time %in% currFit$Time, 
		"prediction", "actual")

# Create a png graph showing the currencies in the context 
# of the Swiss Franc
png(filename = "swiss_franc.png", width = 800, height = 600)
p = ggplot(gcurr, aes(Time, Value, colour = Currency, linetype = Type)) +
  geom_line() + stat_smooth(method = "loess") +
  scale_y_continuous("Change to baseline", labels = percent) +
  labs(title = "Currencies vs Swiss Franc", x = "") +
  theme(axis.title.y = element_text(size = 10, angle = 90, vjust = 0.3))
print(p)
dev.off()

# Define where the image should be placed via a named region;
# let's put the image two columns left to the data starting 
# in the 5th row
createName(wb, name = "graph",
formula = paste(sheet, idx2cref(c(5, ncol(curr) + 2)), sep = "!"))
# Note: idx2cref converts indices (row, col) to Excel cell references

# Put the image created above at the corresponding location
addImage(wb, filename = "swiss_franc.png", name = "graph",
originalSize = TRUE)

saveWorkbook(wb)


