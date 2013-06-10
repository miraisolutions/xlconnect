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
# Wrapper function to write data to an Excel file in one line.
# 
# Author: Thomas Themel, Mirai Solutions GmbH
#
#############################################################################


writeNamedRegionToFile <- function(file, data, name, formula = NA, ..., 
		styleAction = XLC$STYLE_ACTION.XLCONNECT,clearNamedRegions=FALSE) {

  wb <- loadWorkbook(file, create = !file.exists(file))  
  setStyleAction(wb, styleAction)
 
  # clear existing named regions
  existingNames <- getDefinedNames(wb)
  toClear <- intersect(name[clearNamedRegions], existingNames)
  clearNamedRegion(wb, toClear)
  
  if(!is.na(formula)) {
    sheets = extractSheetName(formula)
    createSheet(wb, sheets[sheets != ""])
    createName(wb, name, formula)
  }
  
  writeNamedRegion(wb, data, name, ...)
  
  saveWorkbook(wb)
  invisible(wb)  
}
