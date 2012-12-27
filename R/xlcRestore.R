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
# Restoring objects from an Excel file (that was created via xlcDump)
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

xlcRestore <- function(file = "dump.xlsx", pos = -1, overwrite = FALSE) {
	wb = loadWorkbook(file, create = FALSE)
	sheets = setdiff(getSheets(wb), getOption("XLConnect.Sheet"))
	sapply(sheets, function(obj) {
		if(exists(obj, where = pos) && !overwrite)
			return(FALSE)
		tryCatch({
			data = readWorksheet(wb, sheet = obj, rownames = getOption("XLConnect.RownameCol"))
			assign(obj, data, pos = pos)
			TRUE
		}, error = function(e) {
			warning("Object '", obj, "' has not been restored. Reason: ", conditionMessage(e), call. = FALSE)
			FALSE
		})
	})
}
