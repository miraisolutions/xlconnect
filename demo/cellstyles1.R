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
# Using cellstyles in combination with the 'name prefix' style action
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "cellstyles1.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Set style action to 'name prefix' 
# (default is 'XLConnect' (XLC$"STYLE_ACTION.XLCONNECT"))
#
# This will instruct XLConnect to look for named cellstyles
# (with a specified prefix) when writing headers and columns.
setStyleAction(wb, XLC$"STYLE_ACTION.NAME_PREFIX")
# Set the name prefix for the above style action to 'MyPersonalStyle'
setStyleNamePrefix(wb, "MyPersonalStyle")

# We now create a named cellstyle to be used for the header 
# (column names) of a data.frame
headerCellStyle <- createCellStyle(wb, name = "MyPersonalStyle.Header")
# Specify the cellstyle to use a solid foreground color
setFillPattern(headerCellStyle, fill = XLC$"FILL.SOLID_FOREGROUND")
# Specify the foreground color to be used
setFillForegroundColor(headerCellStyle, color = XLC$"COLOR.LIGHT_CORNFLOWER_BLUE")
# Specify a thick black bottom border
setBorder(headerCellStyle, side = "bottom", type = XLC$"BORDER.THICK", color = XLC$"COLOR.BLACK")

# NOTE: You could define a separate header cellstyle for each column.
# To do this, you could define cellstyles like '<NAME_PREFIX>.Header.<COLUMN_NAME>'
# to reference header cellstyles by column name or '<NAME_PREFIX>.Header.<COLUMN_INDEX>'
# to reference header cellstyles by column index.
#
# The same holds for actual columns. '<NAME_PREFIX>.Column.<COLUMN_NAME>' references
# column cellstyles by column name, '<NAME_PREFIX>.Column.<COLUMN_INDEX>' references
# column cellstyles by column index and '<NAME_PREFIX>.Column.<DATA_TYPE>' references
# column cellstyles by data type (where <DATA_TYPE> is one of {Boolean, DateTime, Numeric,
# String}).

# We now create a named cellstyle to be used for the column named 'wt'
# (as you will see below, we will write the built-in data.frame 'mtcars')
wtColumnCellStyle <- createCellStyle(wb, name = "MyPersonalStyle.Column.wt")
# Specify the cellstyle to use a solid foreground color
setFillPattern(wtColumnCellStyle, fill = XLC$"FILL.SOLID_FOREGROUND")
# Specify the foreground color to be used
setFillForegroundColor(wtColumnCellStyle, color = XLC$"COLOR.LIGHT_ORANGE")

# We now create a named cellstyle to be used for the 3rd column in the data.frame
wtColumnCellStyle <- createCellStyle(wb, name = "MyPersonalStyle.Column.3")
# Specify the cellstyle to use a solid foreground color
setFillPattern(wtColumnCellStyle, fill = XLC$"FILL.SOLID_FOREGROUND")
# Specify the foreground color to be used
setFillForegroundColor(wtColumnCellStyle, color = XLC$"COLOR.LIME")

# INFO: See the documentation for more information on style actions
# and cellstyles!

# Create a sheet named 'mtcars'
createSheet(wb, name = "mtcars")

# Create a named region called 'mtcars' referring to the sheet called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$A$1")

# Write built-in data set 'mtcars' to the above defined named region.
# The style action 'name prefix' will be used when writing the data
# as defined above.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") browseURL(file.path(getwd(), demoExcelFile))
}
