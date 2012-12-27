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
# Testing dumping & restoring objects to/from Excel files
# 
# Author: Martin, Mirai Solutions GmbH
#
#############################################################################

test.dumpAndRestore <- function() {
	if(getOption("FULL.TEST.SUITE")) {
		require(datasets)
		pos = "package:datasets"
	
		# Get all data.frame's from 'package:datasets'
		objs = ls(pos = pos)
		idx = sapply(objs, function(obj) is.data.frame(get(obj, pos = pos)))
		objs = objs[idx]
	
		for(file in c("testDumpAndRestore.xls", "testDumpAndRestore.xlsx")) {
			
			out = xlcDump(objs, file = file, pos = pos, overwrite = TRUE)
			xlcRestore(file = file, pos = globalenv(), overwrite = TRUE)
			
			sapply(names(out)[out], function(obj) {
				data.orig = normalizeDataframe(get(obj, pos = pos))
				data.restored = get(obj)
				checkEquals(data.orig, data.restored, check.attributes = FALSE, check.names = TRUE)
				checkEquals(attr(data.orig, "row.names"), attr(data.restored, "row.names"))
			})
		}
	}
}
