#############################################################################
#
# XLConnect
# Copyright (C) 2010-2025 Mirai Solutions GmbH
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
# Tests around appending data to named regions
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.appendNamedRegion <- function() {
  
  test_overwrite_formula <- function(wb, expected, coords, overwrite = TRUE) {
    appendNamedRegion(wb, mtcars, name = "mtcars_formula", overwriteFormulaCells = overwrite)
    res = readNamedRegion(wb, name = "mtcars_formula")
    checkEquals(res, normalizeDataframe(expected), check.attributes = FALSE, check.names = TRUE)
    checkEquals(getReferenceCoordinatesForName(wb, "mtcars_formula"), coords)
  }
  
  # Create workbooks
  wb.xls <- loadWorkbook(rsrc("resources/testWorkbookAppend.xls"))
  wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookAppend.xlsx"))
  
  refCoord = matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)
  
  # Check that appending data to a named region produces the expected result (*.xls)
  appendNamedRegion(wb.xls, mtcars, name = "mtcars")
  res = readNamedRegion(wb.xls, name = "mtcars")
  checkEquals(res, normalizeDataframe(rbind(mtcars, mtcars)), check.attributes = FALSE, check.names = TRUE)
  checkEquals(getReferenceCoordinatesForName(wb.xls, "mtcars"), refCoord)
  
  # Check that appending data to a named region produces the expected result (*.xlsx)
  appendNamedRegion(wb.xlsx, mtcars, name = "mtcars")
  res = readNamedRegion(wb.xlsx, name = "mtcars")
  checkEquals(res, normalizeDataframe(rbind(mtcars, mtcars)), check.attributes = FALSE, check.names = TRUE)
  checkEquals(getReferenceCoordinatesForName(wb.xlsx, "mtcars"), refCoord)
  
  # Check that trying to append to an non-existing named region throws an error (*.xls)
  checkException(appendNamedRegion(wb.xls, mtcars, name = "doesNotExist"))
  
  # Check that trying to append to an non-existing named region throws an error (*.xlsx)
  checkException(appendNamedRegion(wb.xlsx, mtcars, name = "doesNotExist"))
  
  # Initialize mtcars_mod as is defined with the formula in the carb column (set equal to the gear column)
  mtcars_mod = mtcars
  mtcars_mod$carb = mtcars_mod$gear
  
  # Check that appending data to a named region with a formula below it does not overwrite it if
  # overwriteFormulaCells is set to FALSE (*.xls)
  test_overwrite_formula(wb.xls, rbind(mtcars_mod, mtcars_mod), refCoord, overwrite = FALSE)
  
  # Check that appending data to a named region with a formula below it does not overwrite it if
  # overwriteFormulaCells is set to FALSE (*.xlsx)
  test_overwrite_formula(wb.xlsx, rbind(mtcars_mod, mtcars_mod), refCoord, overwrite = FALSE)
  
  # Re-create workbooks
  wb.xls <- loadWorkbook(rsrc("resources/testWorkbookAppend.xls"))
  wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookAppend.xlsx"))
  
  refCoord = matrix(c(9, 5, 194, 15), ncol = 2, byrow = TRUE)
  refCoordFormula = matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)
  
  # Check that appending data to a named region with a different structure results
  # in the correct bounding box for the re-defined named region (*.xls)
  appendNamedRegion(wb.xls, airquality, name = "mtcars")
  checkEquals(getReferenceCoordinatesForName(wb.xls, "mtcars"), refCoord)
  
  # Check that appending data to a named region with a different structure results
  # in the correct bounding box for the re-defined named region (*.xlsx)
  appendNamedRegion(wb.xlsx, airquality, name = "mtcars")
  checkEquals(getReferenceCoordinatesForName(wb.xlsx, "mtcars"), refCoord)
  
  # Check that appending data to a named region with a formula below it overwrites it if
  # overwriteFormulaCells is set to TRUE (default) (*.xls)
  test_overwrite_formula(wb.xls, rbind(mtcars_mod, mtcars), refCoordFormula)
  
  # Check that appending data to a named region with a formula below it overwrites it if
  # overwriteFormulaCells is set to TRUE (default) (*.xlsx)
  test_overwrite_formula(wb.xlsx, rbind(mtcars_mod, mtcars), refCoordFormula)
}
