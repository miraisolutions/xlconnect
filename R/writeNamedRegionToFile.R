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
# Wrapper function to write data to an Excel file in one line.
# 
# Author: Thomas Themel, Mirai Solutions GmbH
#
#############################################################################


writeNamedRegionToFile <- function(file, data, name, formula=NA, header = TRUE, styleAction = XLC$STYLE_ACTION.XLCONNECT) {
  # if formula is not specified, we expect a workbook with predefined regions
  create.names <- !is.na(formula)
  wb <- loadWorkbook(file,create=create.names)  
  setStyleAction(wb, styleAction)
  
  if(create.names) { 
    # extract "SheetX" from "SheetX!$A$1:$B$2"
    sheetNames <- function(formula) {
      sub("!.*", "", formula[grep("!", formula)])
    }

    createSheet(wb, sheetNames(formula))
    createName(wb, name, formula)
  }
  
  writeNamedRegion(wb, data, name, header);
  
  saveWorkbook(wb)
  invisible(wb)  
}
