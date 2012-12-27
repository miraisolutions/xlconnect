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
# Reading data from a worksheet in an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

## CASE 1: 
## Data starting in top left corner; no other data
## contained on same worksheet

# Read worksheet 'mtcars' (providing no specific area bounds;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars")
# Alternatively: wb$readWorksheet(sheet = "mtcars")

# Print resulting data.frame
print(data)

## CASE 2: 
## Data offset from top left corner; no other data
## contained on same worksheet

# Read worksheet 'mtcars2' (providing no specific area bounds;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars2")
# Alternatively: wb$readWorksheet(sheet = "mtcars2")

# Print resulting data.frame
print(data)

## CASE 3: 
## Data offset from top left corner; also, other data contained
## on same worksheet

# Read worksheet 'mtcars3' (providing specific area bounds;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars3", startRow = 10, startCol = 6,
		endRow = 42, endCol = 16)
# Alternatively: wb$readWorksheet(sheet = "mtcars3", startRow = 10, startCol = 6,
#					endRow = 42, endCol = 16)

# Print resulting data.frame
print(data)
