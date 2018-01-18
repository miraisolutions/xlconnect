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
	
	pwdProtectedFile <- rsrc("resources/testBug61.xlsx")
	
	# Check that openening a password protected file throws an error
	# if no password is specified
	wb <- checkException(loadWorkbook(pwdProtectedFile))
	
	# Check that opening a password protected file throws an error
	# if a wrong password is specified
	b <- checkException(loadWorkbook(pwdProtectedFile, password = "wrong"))
	
	# Check that a password protected file can be openend if the
	# correct password is specified
	wb <- loadWorkbook(pwdProtectedFile, password = "mirai")
}
