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
# Wrapper function to write data to specified locations in  an Excel file in
# one line.
# 
# Author: Thomas Themel, Mirai Solutions GmbH
#
#############################################################################


writeWorksheetToFile <- function(file, styleAction = XLC$STYLE_ACTION.XLCONNECT, ...) {
  args <- list(...)

  # only create sheets when we have names specified
  create <- is.character(args$sheet)

  wb <- loadWorkbook(file,create=create)  
  setStyleAction(wb,styleAction) # new line
  if(create) {	
    createSheet(wb, args$sheet)
  }

  args$object <- wb
  do.call(writeWorksheet, args)
  
  saveWorkbook(wb)
  invisible(wb)  
}

