# Directly referencing named regions in an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

require(XLConnect)

# mydata xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mydata.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Use 'with' in conjunction with a 'workbook' and directly
# reference the containing named regions ('mydata' in this case)
# Note: there's no need to explicitly read the named region
with(wb, {
	coplot(mpg ~ disp | as.factor(cyl), data = mydata,
		panel = panel.smooth, rows = 1)
})
