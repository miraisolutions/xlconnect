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
# Helper function for opening workbooks
# This is to be able to write loadWorkbook(...) rather than 
# new("workbook", filename = ..., create = ...)
#
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

loadWorkbook <- function(filename, create = FALSE) {
  # Run Java AWT in headless mode
  .jcall("java.lang.System", "S", "setProperty", "java.awt.headless", "true")
  
	jTryCatch(new("workbook", filename = path.expand(filename), create = create))
}
