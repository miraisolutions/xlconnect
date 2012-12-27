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
# Adding an image/graph to an Excel file
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
## Write an image in original size matching top left corner of a cell

# Create a sheet named 'switzerland1'
createSheet(wb, name = "switzerland1")
# Alternatively: wb$createSheet(name = "switzerland1")

# Create a named region called 'switzerland1' referring to the sheet called 'switzerland1'
createName(wb, name = "switzerland1", formula = "switzerland1!$C$4")
# Alternatively: wb$createName(name = "switzerland1", formula = "switzerland1!$C$4")

# Write image to the named region created above using the image's original size;
# i.e. the image's top left corner will match the specified cell's top left corner
addImage(wb, filename = demoImageFile, name = "switzerland1", originalSize = TRUE)
# Alternatively: wb$addImage(filename = demoImageFile, name = "switzerland1", originalSize = TRUE)

## CASE 2:
## Write an image to a specified named region

# Create a sheet named 'switzerland2'
createSheet(wb, name = "switzerland2")
# Alternatively: wb$createSheet(name = "switzerland2")

# Create a named region called 'switzerland2' referring to the sheet called 'switzerland2'
createName(wb, name = "switzerland2", formula = "switzerland2!$C$4:$F$9")
# Alternatively: wb$createName(name = "switzerland2", formula = "switzerland2!$C$4:$F$9")

# Write image to the named region area created above - image will be fitted accordingly
addImage(wb, filename = demoImageFile, name = "switzerland2", originalSize = FALSE)
# Alternatively: wb$addImage(filename = demoImageFile, name = "switzerland2", originalSize = FALSE)

## CASE 3:
## Write an R plot to a specified named region
## This example makes use of the 'Tonga Trench Earthquakes' example

# Create a sheet named 'earthquake'
createSheet(wb, name = "earthquake")
# Alternatively: wb$createSheet(name = "earthquake")

# Create a named region called 'earthquake' referring to the sheet called 'earthquake'
createName(wb, name = "earthquake", formula = "earthquake!$B$2")
# Alternatively: wb$createName(name = "earthquake", formula = "earthquake!$B$2")

# Create R plot to a png device
require(lattice)
png(filename = "earthquake.png", width = 800, height = 600)
devAskNewPage(ask = FALSE)

Depth <- equal.count(quakes$depth, number=8, overlap=.1)
xyplot(lat ~ long | Depth, data = quakes)
update(trellis.last.object(),
		strip = strip.custom(strip.names = TRUE, strip.levels = TRUE),
		par.strip.text = list(cex = 0.75),
		aspect = "iso")

dev.off()

# Write image to the named region created above using the image's original size;
# i.e. the image's top left corner will match the specified cell's top left corner
addImage(wb, filename = "earthquake.png", name = "earthquake", originalSize = TRUE)
# Alternatively: wb$addImage(filename = demoImageFile, name = "switzerland1", originalSize = TRUE)


# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
# Alternatively: wb$saveWorkbook()

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") browseURL(file.path(getwd(), demoExcelFile))
}
