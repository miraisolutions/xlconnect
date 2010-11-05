#############################################################################
#
# XLConnect
# Copyright (C) 2010 Mirai Solutions GmbH
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
# Adding an image to an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "image.xlsx"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Image file to write
demoImageFile <- system.file("demoFiles/SwitzerlandFlag.jpg", package = "XLConnect")

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

## CASE 1:
## Write image in original size matching top left corner of a cell and the image

# Create a named region called 'switzerland1' on a sheet called 'switzerland1'
# (the call to 'createName' automatically creates the sheet
# referenced in the formula if it does not exist)
createName(wb, name = "switzerland1", formula = "switzerland1!$C$4")
# Alternatively: wb$createName(name = "switzerland1", formula = "switzerland1!$C$4")

# Write image to the named region created above using the image's original size;
# i.e. the image's top left corner will match the specified cell's top left corner
addImage(wb, filename = demoImageFile, name = "switzerland1", originalSize = TRUE)
# Alternatively: wb$addImage(filename = demoImageFile, name = "switzerland1", originalSize = TRUE)

## CASE 2:
## Write image to a specified named region area

# Create a named region called 'switzerland2' on a sheet called 'switzerland2'
# (the call to 'createName' automatically creates the sheet
# referenced in the formula if it does not exist)
createName(wb, name = "switzerland2", formula = "switzerland2!$C$4:$F$9")
# Alternatively: wb$createName(name = "switzerland2", formula = "switzerland2!$C$4:$F$9")

# Write image to the named region area created above - image will be fitted accordingly
addImage(wb, filename = demoImageFile, name = "switzerland2", originalSize = FALSE)
# Alternatively: wb$addImage(filename = demoImageFile, name = "switzerland2", originalSize = FALSE)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
# Alternatively: wb$saveWorkbook()

if(interactive() && exists("shell.exec")) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") shell.exec(file.path(getwd(), demoExcelFile))
}
