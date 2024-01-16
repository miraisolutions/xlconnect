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
# Tests around setting flags to force excel to recalculate formula values.
# Tests also getting said flag, assuming getForceFormulaRecalculation works
# as intended.
# 
# Author: Peter Schmid, Mirai Solutions GmbH
#
#############################################################################

test.workbook.setForceFormulaRecalculation <- function() {
  
  # Create workbook
  wb.xlsx <- loadWorkbook(rsrc("resources/testBug170.xlsx"), create = FALSE)
  
  # Check that setting the force formula recalculation flag works fine (*.xlsx)
  # (assumes that 'getForceFormulaRecalculation' works fine)
  setForceFormulaRecalculation(wb.xlsx, 1, TRUE)
  checkTrue(getForceFormulaRecalculation(wb.xlsx, 1))
  setForceFormulaRecalculation(wb.xlsx, c("Sheet1", "Sheet2"), FALSE)
  checkTrue(!getForceFormulaRecalculation(wb.xlsx, "Sheet2"))
  
  # Check that passing multiple sheets doesn't cause problems
  setForceFormulaRecalculation(wb.xlsx, "*", TRUE)
  checkTrue(all(getForceFormulaRecalculation(wb.xlsx, "*")))
  
  # Check that setting the force formula recalculation flag an illegal active sheet throws an exception (*.xlsx)
  checkException(setForceFormulaRecalculation(wb.xlsx, 12, TRUE))
  checkException(setForceFormulaRecalculation(wb.xlsx, "SheetWhichDoesNotExist", TRUE))
}
