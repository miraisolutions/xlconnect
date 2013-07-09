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
# General XLConnect Settings
# Called by .onLoad which is also passing the package description (pdesc)
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

XLConnectSettings <- function(pdesc) {
	
	# URL to Mirai Solutions GmbH Website
	options(MiraiSolutions.URL = pdesc$URL)
	
	# XLConnect Version
	options(XLConnect.Version = pdesc$Version)
	
	# Date/time format used for conversion to string;
	# This is used for communicating a string-based date/time
	# representation to Java which will then convert it to java.util.Date
	options(XLConnect.dateTimeFormat = "%Y-%m-%d %H:%M:%S")
	
	# Dummy sheet name (used by xlcDump, xlcRestore)
	options(XLConnect.Sheet = "#xlc#")
	# Rowname column (used by xlcDump, xlcRestore)
	options(XLConnect.RownameCol = ".rownames")
	
	invisible()
}
