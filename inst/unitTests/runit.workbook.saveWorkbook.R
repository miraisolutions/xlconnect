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
# Tests around saving Excel workbooks
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.saveWorkbook <- function() {
  if(getOption("FULL.TEST.SUITE")) {
  	# Create workbooks
  	file.xls <- rsrc("resources/testWorkbookSaveWorkbook.xls")
  	file.xlsx <- rsrc("resources/testWorkbookSaveWorkbook.xlsx")
  	file.remove(file.xls)
  	file.remove(file.xlsx)
  	wb.xls <- loadWorkbook(file.xls, create = TRUE)
  	wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
  	
  	# Files don't exist yet
  	checkTrue(!file.exists(file.xls))
  	checkTrue(!file.exists(file.xlsx))
  	
  	saveWorkbook(wb.xls)
  	saveWorkbook(wb.xlsx)
  	
  	# Check that file exists after saving (*.xls)
  	checkTrue(file.exists(file.xls))
  	
  	# Check that file exists after saving (*.xlsx)
  	checkTrue(file.exists(file.xlsx))
  	
  	# Check save as (*.xls)
  	newFile.xls <- "saveAsWorkbook.xls"
  	if(file.exists(newFile.xls)) file.remove(newFile.xls)
  	saveWorkbook(wb.xls, file = newFile.xls)
  	checkTrue(file.exists(newFile.xls))
  	
  	# Check save as (*.xlsx)
  	newFile.xlsx <- "saveAsWorkbook.xlsx"
  	if(file.exists(newFile.xlsx)) file.remove(newFile.xlsx)
  	saveWorkbook(wb.xlsx, file = newFile.xlsx)
  	checkTrue(file.exists(newFile.xlsx))
  }
}
