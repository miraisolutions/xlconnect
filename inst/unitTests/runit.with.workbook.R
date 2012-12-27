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
# Tests around using Excel workbooks in a with(...) context
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.with.workbook <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWithWorkbook.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWithWorkbook.xlsx"), create = FALSE)
	
	# Check if named regions can be correctly referenced (*.xls)
	with(wb.xls, {
		checkTrue(all(dim(AA) == c(8, 3)))
		checkTrue(all(dim(BB) == c(5, 5)))
	}, header = FALSE)

	# Check if named regions can be correctly referenced (*.xlsx)
	with(wb.xlsx, {
		checkTrue(all(dim(AA) == c(8, 3)))
		checkTrue(all(dim(BB) == c(5, 5)))
	}, header = FALSE)
}
