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
# Dumping objects to an Excel file
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

xlcDump <- function(list, ..., file = "dump.xlsx", pos = -1, overwrite = FALSE) {
	# Name of dummy sheet
	# (required since there always needs to be at least one worksheet)
	xlc = getOption("XLConnect.Sheet")

	wb = loadWorkbook(file, create = TRUE)
	if(missing(list))
		list = ls(..., pos = pos)
	# Create dummy sheet
	createSheet(wb, name = xlc)
	sheets = setdiff(getSheets(wb), xlc)
	
	res = sapply(list, function(obj) {
		sin = obj %in% sheets
		if(sin && overwrite)
			removeSheet(wb, sheet = obj)
		else if(sin && !overwrite)
			return(FALSE)
		
		createSheet(wb, name = obj)
		tryCatch({
			writeWorksheet(wb, data = get(obj, pos = pos), sheet = obj,
				rownames = getOption("XLConnect.RownameCol"))
			TRUE
		}, error = function(e) {
			removeSheet(wb, sheet = obj)
			warning("Object '", obj, "' has not been written. Reason: ", conditionMessage(e), call. = FALSE)
			FALSE
		})
	})
	# Remove dummy sheet if there is at least one other
	if(length(getSheets(wb)) > 1)
		removeSheet(wb, sheet = xlc)
	saveWorkbook(wb)
	res
}
