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
# Tests around cell styles
# 
# Author: Thomas Themel, Mirai Solutions GmbH
#
#############################################################################

test.cellStyle <- function() {
        # Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookCellStyles.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookCellStyles.xlsx"), create = TRUE)

        # It should not contain a style
        checkException(getCellStyle(wb.xls, 'Header'))
        checkException(getCellStyle(wb.xlsx, 'Header'))

        # Create some styles
        cs.xls <- createCellStyle(wb.xls, 'Header')
        cs.xlsx <- createCellStyle(wb.xlsx, 'Header')

        # Now we should be able to get them.
        # NB: cs.xls and cs.xls2 refer to different objects!        
        cs.xls2 <- getCellStyle(wb.xls, 'Header')
        cs.xlsx2 <- getCellStyle(wb.xlsx, 'Header')
}
