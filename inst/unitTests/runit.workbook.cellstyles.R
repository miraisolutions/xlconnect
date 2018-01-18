#############################################################################
#
# XLConnect
# Copyright (C) 2010-2018 Mirai Solutions GmbH
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
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Tests around cell styles
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.cellstyles <- function() {
  if(getOption("FULL.TEST.SUITE")) {
    
    file.xls <- rsrc("resources/cellstyles.xls")
    file.xlsx <- rsrc("resources/cellstyles.xlsx")
    file.remove(file.xls)
    file.remove(file.xlsx)
    
    # Create workbooks
    wb.xls <- loadWorkbook(file.xls, create = TRUE)
    wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
    
    styleName <- "MyStyle"
    anotherStyleName <- "MyOtherStyle"
    
    # Check that a cell style which hasn't been created yet does not exist
    checkTrue(!existsCellStyle(wb.xls, styleName))
    checkTrue(!existsCellStyle(wb.xlsx, styleName))
    
    # Using getCellStyle is expected to through an exception
    checkException(getCellStyle(wb.xls, styleName))
    checkException(getCellStyle(wb.xlsx, styleName))
    
    # Check that a cell style can be created
    checkNoException(createCellStyle(wb.xls, styleName))
    checkNoException(createCellStyle(wb.xlsx, styleName))
    
    # Check that a cell style which has been created exists
    checkTrue(existsCellStyle(wb.xls, styleName))
    checkTrue(existsCellStyle(wb.xlsx, styleName))
    
    # Attempting to create a cell style which already exists is expected
    # to throw an exception
    checkException(createCellStyle(wb.xls, styleName))
    checkException(createCellStyle(wb.xlsx, styleName))
    
    # Check that a cell style which has been created can be retrieved
    checkNoException(getCellStyle(wb.xls, styleName))
    checkNoException(getCellStyle(wb.xlsx, styleName))
    
    # Check creation and retrieval of cell styles using getOrCreateCellStyle
    checkTrue(!existsCellStyle(wb.xls, anotherStyleName))
    checkTrue(!existsCellStyle(wb.xlsx, anotherStyleName))
    checkNoException(getOrCreateCellStyle(wb.xls, anotherStyleName))
    checkNoException(getOrCreateCellStyle(wb.xlsx, anotherStyleName))
    checkTrue(existsCellStyle(wb.xls, anotherStyleName))
    checkTrue(existsCellStyle(wb.xlsx, anotherStyleName))
    checkNoException(getOrCreateCellStyle(wb.xls, anotherStyleName))
    checkNoException(getOrCreateCellStyle(wb.xlsx, anotherStyleName))
  }
}
