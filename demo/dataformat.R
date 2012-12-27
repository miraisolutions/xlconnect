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
# Using Excel data formats in combination with the 'data format only' 
# style action.
# This is useful when writing data to an Excel template where all styling
# should be kept intact and only the data format for a cell should be set.
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "dataformat.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)
# Copy the existing Excel template.
# This template contains a sheet 'mtcars' and defines a named region 'mtcars'
# that is already pre-styled with some custom cellstyles for headers and columns.
# You may have a look at this file to compare it with the final result.
file.copy(system.file("demoFiles/template2.xlsx", package = "XLConnect"), demoExcelFile)

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Set the data format for numeric columns (cells)
# (keeping the defaults for all other data types)
setDataFormatForType(wb, type = XLC$"DATA_TYPE.NUMERIC", format = "0.00")

# Set style action to 'data format only' 
# (default is 'XLConnect' (XLC$"STYLE_ACTION.XLCONNECT"))
#
# This will instruct XLConnect to only apply data formats
# when writing headers and columns.
setStyleAction(wb, XLC$"STYLE_ACTION.DATA_FORMAT_ONLY")

# INFO: See the documentation for more information on style actions
# and cellstyles!

# Write built-in data set 'mtcars' to the named region 'mtcars' as defined
# by the Excel template.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") browseURL(file.path(getwd(), demoExcelFile))
}
