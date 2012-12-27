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
# Directly referencing named regions in an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

require(XLConnect)

# mydata xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mydata.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Use 'with' in conjunction with a 'workbook' and directly
# reference the containing named regions ('mydata' in this case)
# Note: there's no need to explicitly read the named region

with(wb, {
	coplot(mpg ~ disp | as.factor(cyl), data = mydata,
		panel = panel.smooth, rows = 1)
})


# CAUTION: 'with' should only be used with simple workbooks that
# contain a small number of named regions; if you have workbooks
# with many named regions and you are only interested in a particluar
# subset, consider reading in those named region "manually" as this
# may increase performance
