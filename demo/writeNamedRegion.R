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
# Writing a named region to an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "mtcars.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Create a worksheet named 'mtcars'
createSheet(wb, name = "mtcars")
# Alternatively: wb$createSheet(name = "mtcars")

# Create a named region called 'mtcars' on the sheet called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$A$1")
# Alternatively: wb$createName(name = "mtcars", formula = "mtcars!$A$1")

# Write built-in data set 'mtcars' to the above defined named region
writeNamedRegion(wb, mtcars, name = "mtcars")
# Alternatively: wb$writeNamedRegion(mtcars, name = "mtcars")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
# Alternatively: wb$saveWorkbook()

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") browseURL(file.path(getwd(), demoExcelFile))
}
