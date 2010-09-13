# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.createName <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/createName.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/createName.xlsx"), create = TRUE)
	
	# Check that creating a legal name works ok (*.xls)
	# (this test assumes 'existsName' is working fine)
	createName(wb.xls, "Test", "Test!$C$5")
	checkTrue(existsName(wb.xls, "Test"))
	
	# Check that creating a legal name works ok (*.xlsx)
	# (this test assumes 'existsName' is working fine)
	createName(wb.xlsx, "Test", "Test!$C$5")
	checkTrue(existsName(wb.xlsx, "Test"))
	
	# Check that trying to create an illegal name throws
	# an exception (*.xls)
	checkException(createName(wb.xls, "'Test", "Test!$C$10"))
	
	# Check that trying to create an illegal name throws
	# an exception (*.xlsx)
	checkException(createName(wb.xlsx, "'Test", "Test!$C$10"))
	
	# Check that trying to create a name with an illegal formula
	# throws an exception (*.xls)
	checkException(createName(wb.xls, "IllegalFormula", "öö-%&"))
	
	# Check that trying to create a name with an illegal formula
	# throws an exception (*.xlsx)
	checkException(createName(wb.xlsx, "IllegalFormula", "öö-%&"))
	
	# Check that trying to create an already existing name without
	# specifying 'overwrite = TRUE' throws an exception (*.xls)
	createName(wb.xls, "ImHere", "ImHere!$B$9")
	checkException(createName(wb.xls, "ImHere", "There!$A$2"))
	
	# Check that trying to create an already existing name without
	# specifying 'overwrite = TRUE' throws an exception (*.xlsx)
	createName(wb.xlsx, "ImHere", "ImHere!$B$9")
	checkException(createName(wb.xlsx, "ImHere", "There!$A$2"))
	
	# Check that overwriting an existing name works ok (*.xls)
	createName(wb.xls, "CurrentlyHere", "CurrentlyHere!$D$8")
	createName(wb.xls, "CurrentlyHere", "NowThere!$C$3", overwrite = TRUE)
	# TODO: Should actually rather check that new formula is correct
	checkTrue(existsName(wb.xls, "CurrentlyHere"))
	
	# Check that overwriting an existing name works ok (*.xlsx)
	createName(wb.xlsx, "CurrentlyHere", "CurrentlyHere!$D$8")
	createName(wb.xlsx, "CurrentlyHere", "NowThere!$C$3", overwrite = TRUE)
	# TODO: Should actually rather check that new formula is correct
	checkTrue(existsName(wb.xlsx, "CurrentlyHere"))
}

