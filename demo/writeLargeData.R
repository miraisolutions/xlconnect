# Writing large datasets to Excel
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

# Increase Java's maximum heap space;
# this option must be set before the underlying JVM is initialized
# and therefore MUST happen before XLConnect is loaded
options(java.parameters = "-Xmx1024m")

if(any(is.element(c("package:XLConnect", "package:rJava"), search()))) {
	msg <- paste(
		"XLConnect and/or rJava are already attached.",
		"You may reload these packages in order for the Java parameter setting",
		"to take effect.",
		sep = "\n"
	)
	warning(msg)
}

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "large.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Create a large dummy data.frame
set.seed(292)
n <- 50000
dfLarge <- data.frame(
	A = rnorm(n),
	B = sample(letters, size = n, replace = TRUE),
	C = rnorm(n) > 0,
	D = rep(as.Date("2010-09-17"), n),
	E = rnorm(n),
	F = sample(LETTERS, size = n, replace = TRUE),
	G = 1:n
)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Create a worksheet called 'Dummy'
createSheet(wb, name = "Dummy")

# Write large data.frame to worksheet 'Dummy' created above
writeWorksheet(wb, dfLarge, sheet = "Dummy")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

if(interactive() && exists("shell.exec")) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") shell.exec(file.path(getwd(), demoExcelFile))
}
