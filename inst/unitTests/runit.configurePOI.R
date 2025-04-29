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
# Test POI configuration
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.configurePOI <- function() {
  
  # Check that an exception is thrown if we limit the number of files to just 1
  configurePOI(zip_max_files = 1L)
  checkException(loadWorkbook(rsrc("resources/testLoadWorkbook.xlsx")))
  
  # Check zip bomb detection with high inflate ratio
  configurePOI(zip_min_inflate_ratio = 0.99)
  checkException(readWorksheetFromFile(
    rsrc("resources/testZipBomb.xlsx"),
    sheet = 1
  ))
  
  # Check that an exception is thrown if we limit the max zip entry size to just
  # 1 byte
  configurePOI(zip_max_entry_size = 1L)
  checkException(readWorksheetFromFile(
    rsrc("resources/testWorkbookReadWorksheet.xlsx"),
    sheet = 1
  ))
  
  # Check storing of zip entries in temp files
  configurePOI(zip_entry_threshold_bytes = 0L)
  readWorksheetFromFile(
    rsrc("resources/testWorkbookReadWorksheet.xlsx"),
    sheet = 1
  )
  
  # Reset settings after test
  configurePOI()
  
}
