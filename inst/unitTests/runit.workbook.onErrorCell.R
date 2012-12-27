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
# Testing different behaviors with error cells
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.onErrorCell <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookErrorCell.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookErrorCell.xlsx"), create = FALSE)
	
	# Check that reading error cells with the warning flag set does not cause any issues (*.xls)
	onErrorCell(wb.xls, XLC$ERROR.WARN)
	
	res <- try(readNamedRegion(wb.xls, name = "AA"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(A = c("aa", "bb", "cc", NA, "ee", "ff"), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	res <- try(readNamedRegion(wb.xls, name = "BB"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(B = c(4.3, NA, -2.5, 1.6, NA, 9.7), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	res <- try(readNamedRegion(wb.xls, name = "CC"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(C = c(-53.2, NA, 34.1, -37.89, 0, 1.6), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	res <- try(readNamedRegion(wb.xls, name = "DD"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(D = c(8.2, 2, 1, -0.5, NA, 3.1), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	res <- try(readNamedRegion(wb.xls, name = "EE"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(E = c("zz", "yy", NA, "ww", "vv", "uu"), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	# Check that reading error cells with the warning flag set does not cause any issues (*.xlsx)
	onErrorCell(wb.xlsx, XLC$ERROR.WARN)
	
	res <- try(readNamedRegion(wb.xlsx, name = "AA"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(A = c("aa", "bb", "cc", NA, "ee", "ff"), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	res <- try(readNamedRegion(wb.xlsx, name = "BB"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(B = c(4.3, NA, -2.5, 1.6, NA, 9.7), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	res <- try(readNamedRegion(wb.xlsx, name = "CC"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(C = c(-53.2, NA, 34.1, -37.89, 0, 1.6), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	res <- try(readNamedRegion(wb.xlsx, name = "DD"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(D = c(8.2, 2, 1, -0.5, NA, 3.1), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	res <- try(readNamedRegion(wb.xlsx, name = "EE"))
	checkTrue(!is(res, "try-error"))
	target <- data.frame(E = c("zz", "yy", NA, "ww", "vv", "uu"), stringsAsFactors = FALSE)
	checkEquals(res, target)
	
	# Check that reading error cells with the stop flag set causes an exception (*.xls)
	onErrorCell(wb.xls, XLC$ERROR.STOP)
	
	res <- try(readNamedRegion(wb.xls, name = "AA"))
	checkTrue(is(res, "try-error"))
	
	res <- try(readNamedRegion(wb.xls, name = "BB"))
	checkTrue(is(res, "try-error"))
	
	res <- try(readNamedRegion(wb.xls, name = "CC"))
	checkTrue(is(res, "try-error"))
	
	res <- try(readNamedRegion(wb.xls, name = "DD"))
	checkTrue(is(res, "try-error"))
	
	res <- try(readNamedRegion(wb.xls, name = "EE"))
	checkTrue(is(res, "try-error"))
	
	# Check that reading error cells with the stop flag set causes an exception (*.xlsx)
	onErrorCell(wb.xlsx, XLC$ERROR.STOP)
	
	res <- try(readNamedRegion(wb.xlsx, name = "AA"))
	checkTrue(is(res, "try-error"))
	
	res <- try(readNamedRegion(wb.xlsx, name = "BB"))
	checkTrue(is(res, "try-error"))
	
	res <- try(readNamedRegion(wb.xlsx, name = "CC"))
	checkTrue(is(res, "try-error"))
	
	res <- try(readNamedRegion(wb.xlsx, name = "DD"))
	checkTrue(is(res, "try-error"))
	
	res <- try(readNamedRegion(wb.xlsx, name = "EE"))
	checkTrue(is(res, "try-error"))
}
