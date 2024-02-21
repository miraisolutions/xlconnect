#############################################################################
#
# XLConnect
# Copyright (C) 2010-2024 Mirai Solutions GmbH
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Tests around creating Excel names
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.createName <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/createName.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/createName.xlsx"), create = TRUE)
	
	# Check that creating a legal name works ok (*.xls)
	# (this test assumes 'existsName' is working fine)
	createName(wb.xls, "Test", "Test!$C$5")
	checkEquals(existsName(wb.xls, "Test"), TRUE, check.attributes = FALSE)
	
	# Check that creating a legal name works ok (*.xlsx)
	# (this test assumes 'existsName' is working fine)
	createName(wb.xlsx, "Test", "Test!$C$5")
	checkEquals(existsName(wb.xlsx, "Test"), TRUE, check.attributes = FALSE)
	
	# Check that trying to create an illegal name throws
	# an exception (*.xls)
	checkException(createName(wb.xls, "'Test", "Test!$C$10"))
	
	# Check that trying to create an illegal name throws
	# an exception (*.xlsx)
	checkException(createName(wb.xlsx, "'Test", "Test!$C$10"))
	
	# Check that trying to create a name with an illegal formula
	# throws an exception (*.xls)
	checkException(createName(wb.xls, "IllegalFormula", "??-%&"))
	
	# Check that trying to create a name with an illegal formula
	# throws an exception (*.xlsx)
	checkException(createName(wb.xlsx, "IllegalFormula", "??-%&"))
	
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
	checkEquals(existsName(wb.xls, "CurrentlyHere"), TRUE, check.attributes = FALSE)
	
	# Check that overwriting an existing name works ok (*.xlsx)
	createName(wb.xlsx, "CurrentlyHere", "CurrentlyHere!$D$8")
	createName(wb.xlsx, "CurrentlyHere", "NowThere!$C$3", overwrite = TRUE)
	# TODO: Should actually rather check that new formula is correct
	checkEquals(existsName(wb.xlsx, "CurrentlyHere"), TRUE, check.attributes = FALSE)
	
	# Check that after trying to write a name with an illegal formula
	# (which throws an exception), the name remains available (*.xls)
	checkException(createName(wb.xls, "aName", "Test!A1A4"))
	checkNoException(createName(wb.xls, "aName", "Test!A1"))
	checkEquals(existsName(wb.xls, "aName"), TRUE, check.attributes = FALSE)
	
	# Check that after trying to write a name with an illegal formula
	# (which throws an exception), the name remains available (*.xlsx)
  #
  # NOTE: This seems to have changed with POI 3.11-beta1 - creating a
  # name with an invalid formula does not throw an exception anymore!
	# checkException(createName(wb.xlsx, "aName", "Test!A1A4"))
	# checkNoException(createName(wb.xlsx, "aName", "Test!A1"))
	# checkEquals(existsName(wb.xlsx, "aName"), TRUE, check.attributes = FALSE)


	# Check that creating a worksheet-scoped name works (*.xls)
	# we first need to create a worksheet:
	# global names can be created without an existing sheet, but not scoped ones.
	
	sheetName <- "Test_Scoped_Sheet"

	createSheet(wb.xls, name = sheetName)
	createSheet(wb.xlsx, name = sheetName)

	checkTrue(existsSheet(wb.xls, sheetName))
	checkTrue(existsSheet(wb.xlsx, sheetName))

	expect_found = TRUE
	attributes(expect_found) <- list(worksheetScope = sheetName)
	createName(wb.xls, "Test_1", "Test!$C$5", worksheetScope = sheetName)
	checkEquals(existsName(wb.xls, "Test_1"), expect_found)

	# Check that creating a worksheet-scoped name works (*.xlsx)
	# (this test assumes 'existsName' is working fine)
	expect_found = TRUE
	attributes(expect_found) <- list(worksheetScope = sheetName)
	createName(wb.xlsx, "Test_1", "Test!$C$5", worksheetScope = sheetName)
	checkEquals(existsName(wb.xlsx, "Test_1"), expect_found)
}

