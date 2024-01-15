#############################################################################
#
# XLConnect
# Copyright (C) 2010-2024 Mirai Solutions GmbH
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
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Creates a Java sequence length encoding object for efficient encoding
# of sequences
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

.jsle <- function(x) {
  sle <- seqle(as.integer(x))
  new(
    J("com.miraisolutions.xlconnect.utils.SequenceLengthEncoding"), 
    .jarray(sle$values),
    .jarray(sle$lengths),
    as.integer(sle$increment)
  )
}
