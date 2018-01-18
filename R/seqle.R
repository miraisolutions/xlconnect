#############################################################################
#
# XLConnect
# Copyright (C) 2010-2018 Mirai Solutions GmbH
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
# Sequence length encoding (based on base::rle)
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

seqle <- function(x) {
  n <- length(x)
  
  # Determine reasonable sequence increment
  d <- rle(sort(diff(x)))
  inc <- d$values[which.max(d$lengths)[1]]
  inc <- ifelse(is.na(inc), 1, inc)
  
  y <- x[-1L] != x[-n] + inc
  i <- c(which(y | is.na(y)), n)
  list(lengths = diff(c(0L, i)), values = x[head(c(0L, i) + 1L, -1L)], increment = inc)
}
