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
# Tests around hiding Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.hideSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHideSheet.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHideSheet.xlsx"), create = TRUE)
	
	# Sheet names
	visibleSheet <- "VisibleSheet"
	hiddenSheet <- "HiddenSheet"
	veryHiddenSheet <- "VeryHiddenSheet"

	# Create some sheets
	createSheet(wb.xls, visibleSheet)
	createSheet(wb.xlsx, visibleSheet)
	createSheet(wb.xls, hiddenSheet)
	createSheet(wb.xlsx, hiddenSheet)
	createSheet(wb.xls, veryHiddenSheet)
	createSheet(wb.xlsx, veryHiddenSheet)
	
	# Check if sheets are hidden correspondingly (*.xls)
	hideSheet(wb.xls, hiddenSheet)
	checkTrue(isSheetHidden(wb.xls, hiddenSheet))
	hideSheet(wb.xls, veryHiddenSheet, veryHidden = TRUE)
	checkTrue(isSheetVeryHidden(wb.xls, veryHiddenSheet))
	
	# Check if sheets are hidden correspondingly (*.xls)
	hideSheet(wb.xlsx, hiddenSheet)
	checkTrue(isSheetHidden(wb.xlsx, hiddenSheet))
	hideSheet(wb.xlsx, veryHiddenSheet, veryHidden = TRUE)
	checkTrue(isSheetVeryHidden(wb.xlsx, veryHiddenSheet))
	
}
