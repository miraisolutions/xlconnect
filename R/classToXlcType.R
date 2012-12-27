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
# Helper function to map R classes to XLC types (XLC$DATA_TYPE.?)
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

classToXlcType <- function(cls) {
	if(length(cls) > 0) {
		xlcTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.DATETIME)
		as.vector(sapply(cls, function(x) {
			switch(x,
				"numeric" = XLC$DATA_TYPE.NUMERIC,
				"integer" = XLC$DATA_TYPE.NUMERIC,
				"logical" = XLC$DATA_TYPE.BOOLEAN,
				"character" = XLC$DATA_TYPE.STRING,
				"Date" = XLC$DATA_TYPE.DATETIME,
				"POSIXt" = XLC$DATA_TYPE.DATETIME,
				"POSIXct" = XLC$DATA_TYPE.DATETIME,
				{
					if(x %in% xlcTypes)
						x
					else {
						warning(sprintf("Class or XLConnect data type '%s' is not supported! Continue with 'character'.", x),
							call. = FALSE)
						XLC$DATA_TYPE.STRING
					}
				}
			)
		}))
	} else
		cls
}
