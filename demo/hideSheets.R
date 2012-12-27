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
# Hiding worksheets of an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# Excel workbook to write
demoExcelFile <- "hide.xls"

# Remove file if it already exists
if(file.exists(demoExcelFile)) file.remove(demoExcelFile)

# Load workbook (create if not existing)
wb <- loadWorkbook(demoExcelFile, create = TRUE)

# Write a couple of built-in data.frame's into sheets
# with corresponding name
for(obj in c("CO2", "airquality", "swiss")) {
	createSheet(wb, name = obj)
	writeWorksheet(wb, get(obj), sheet = obj)
}

# Hide sheet 'airquality';
# the sheet may be unhidden by a user from within Excel
hideSheet(wb, sheet = "airquality")
# Alternatively: wb$hideSheet(sheet = "airquality")

# Very-hide sheet 'swiss';
# the sheet cannot be unhidden by a user from within Excel
hideSheet(wb, sheet = "swiss", veryHidden = TRUE)
# Alternatively: wb$hideSheet(sheet = "swiss", veryHidden = TRUE)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
# Alternatively: wb$saveWorkbook()

if(interactive()) {
	answer <- readline("Open the created Excel file (y/n)? ")
	if(answer == "y") browseURL(file.path(getwd(), demoExcelFile))
}
