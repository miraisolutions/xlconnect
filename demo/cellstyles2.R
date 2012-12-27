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
# Using cellstyles in combination with the 'predefined' style action.
# This is useful when writing data to an Excel template.
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "cellstyles2.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)
# Copy the existing Excel template.
# This template contains a sheet 'mtcars' and defines a named region 'mtcars'
# that is already pre-styled with some custom cellstyles for headers and columns.
# You may have a look at this file to compare it with the final result.
file.copy(system.file("demoFiles/template.xlsx", package = "XLConnect"), demoExcelFile)

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Set style action to 'predefined' 
# (default is 'XLConnect' (XLC$"STYLE_ACTION.XLCONNECT"))
#
# This will instruct XLConnect to use existing (predefined) cellstyles
# when writing headers and columns.
setStyleAction(wb, XLC$"STYLE_ACTION.PREDEFINED")

# INFO: See the documentation for more information on style actions
# and cellstyles!

# Write built-in data set 'mtcars' to the named region 'mtcars' as defined
# by the Excel template.
# Predefined cellstyles will be used as defined by the Excel template
# when writing the data.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") browseURL(file.path(getwd(), demoExcelFile))
}
