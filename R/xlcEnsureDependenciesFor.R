#############################################################################
#
# XLConnect
# Copyright (C) 2010-2020 Mirai Solutions GmbH
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
# XLConnect Package Installation: check and download dependencies if missing
# 
# Author: Simon Poltier, Mirai Solutions GmbH
#
#############################################################################

xlcEnsureDependenciesFor <- function (url, name, pattern, ifFoundPath, libname, pkgname) {
  numJars = length(list.files("/usr/share/java/", pattern = pattern))
  if(numJars == 0) {
    sharedPaths <- c()
    if (!interactive()) {
      destDir <- file.path(libname, pkgname, "java")
      dst <- file.path(destDir, name)
      if(!file.exists(dst)){
        download.file(url, dst, mode="wb")
      }
    }
  } else {
    sharedPaths <- c(ifFoundPath)
  }
  sharedPaths
}
