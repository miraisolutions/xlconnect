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
# Test reproducing error log from issue 181 / PR 183
# 
# Author: Peter Schmid, Mirai Solutions GmbH
#
#############################################################################

test.workbook.testBug181 <- function() {
  # loading this workbook reproduces this error, but only the first time it is loaded during a session
  # it seems however that this does not prevent the workbook from being read and even when the "error"
  # is logged it isn't considered an exception and behaves more like a regular log message
  checkNoException({
    wb <- loadWorkbook(rsrc("resources/testBug181.xlsx"))
    obj <- readWorksheet(wb, "Sheet1")
    checkTrue(is.data.frame(obj))
  })
}
