#############################################################################
#
# XLConnect
# Copyright (C) 2010-2013 Mirai Solutions GmbH
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
# Tests around opening an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.loadWorkbook <- function() {
	
	# Check that an exception is thrown when trying to open
	# a non-existent file (*.xls)
	checkException(loadWorkbook(rsrc("resources/fileWhichDoesNotExist.xls")))
	
	# Check that an exception is thrown when trying to open
	# a non-existent file (*.xls)
	checkException(loadWorkbook(rsrc("resources/fileWhichDoesNotExist.xlsx")))
	
	# Check that an instance of the workbook class is returned
	# for an already existing file (*.xls)
	wb <- loadWorkbook(rsrc("resources/testLoadWorkbook.xls"))
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for an already existing file (*.xlsx)
	wb <- loadWorkbook(rsrc("resources/testLoadWorkbook.xlsx"))
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for a file created on-the-fly (*.xls)
	wb <- loadWorkbook(rsrc("resources/fileCreatedOnTheFly.xls"), create = TRUE)
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for a file created on-the-fly (*.xlsx)
	wb <- loadWorkbook(rsrc("resources/fileCreatedOnTheFly.xlsx"), create = TRUE)
	checkTrue(is(wb, "workbook"))
}

