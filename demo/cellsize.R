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
# Setting row height and column width in a worksheet
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "cellsize.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Prepare the data ...
data <- data.frame(Text = 
			c("Some very very very very long text that we want to fit in a cell",
				"Again, some very very very very long text that we want to fit in a cell"),
			stringsAsFactors = F)

# We now create a (unnamed) cellstyle to be used for wrapping text in a cell
wrapStyle <- createCellStyle(wb)
# Specify the cellstyle to wrap text in a cell
setWrapText(wrapStyle, wrap = TRUE)

# Create a worksheet called 'cellsize'
createSheet(wb, name = "cellsize")

# Write the data set to the worksheet created above;
# offset from the top left corner and with default header = TRUE
writeWorksheet(wb, data, sheet = "cellsize", startRow = 4, startCol = 2)

# Set the wrapStyle cellstyle for the long text cells.
# Note: the row and col arguments are vectorized!
setCellStyle(wb, sheet = "cellsize", row = 5:6, col = 2, cellstyle = wrapStyle)

# Set column width (measured in units of 1/256th of a character width).
# Note: The column and width arguments are vectorized.
setColumnWidth(wb, sheet = "cellsize", column = 2, width = 6000)
# Set row height for data cells (measured in points; -1 sets to default height).
# Note: The row and height arguments are vectorized.
setRowHeight(wb, sheet = "cellsize", row = 5:6, height = 60)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") browseURL(file.path(getwd(), demoExcelFile))
}
