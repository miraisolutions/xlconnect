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
# Utility function to extract the sheet name from a formula
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

extractSheetName <- function(formula) {
	# Note: This matches up to the last occurrence of '!' since the regexpr is greedy
	idx = regexpr(".*!", formula)
	# Extract sheet name
	sheet = substring(formula, idx, idx + attr(idx, "match.length") - 2)
	# Replace leading & trailing '
	gsub("^'|'$", "", sheet)
}
