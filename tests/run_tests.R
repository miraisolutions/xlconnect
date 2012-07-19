#############################################################################
#
# XLConnect
# Copyright (C) 2010-2012 Mirai Solutions GmbH
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
# RUnit test-runner script
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

# Limit number of GC threads
options(java.parameters = c("-XX:+UseParallelGC", "-XX:ParallelGCThreads=1"))

# Option to determine if full test suite should be run
options("FULL.TEST.SUITE" = Sys.getenv("FULL_TEST_SUITE") == "1")

# Load library built by R CMD check
library(package = "XLConnect", character.only = TRUE)
# Run unit tests
runUnitTests()
