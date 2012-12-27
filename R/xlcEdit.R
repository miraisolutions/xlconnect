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
# Function for editing R data.frames in an Excel editor (e.g. MS Excel)
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

xlcEdit <- function(obj, pos = globalenv(), ext = ".xlsx") {
	tmp = tempfile(fileext = ext)
	on.exit(try(unlink(tmp)))
	xlcDump(deparse(substitute(obj)), file = tmp, pos = pos)
	cmd = switch(Sys.info()["sysname"],
		Windows	= 'cmd /C start /WAIT %s',
		Darwin = 'open -W -F %s',
		# Linux	= 'xdg-open %s',
		stop("xlcEdit is currently not supported on your system!")
	)
	system(sprintf(cmd, tmp), wait = TRUE)
	invisible(xlcRestore(file = tmp, overwrite = TRUE, pos = pos))
}
