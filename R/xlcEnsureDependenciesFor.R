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

xlcEnsureDependenciesFor <- function (depSet, pattern, ifFoundPaths, libname, pkgname) {
  numJars = length(list.files("/usr/share/java/", pattern = pattern))
  numDoc = length(list.files("/usr/share/doc/", pattern = pattern))
  if(numJars + numDoc == 0) {
    sharedPaths <- c()
    if (!interactive()) {
      destDir <- file.path(libname, pkgname, "java")
      
      dPairJar <- function (urlAndName) {
        dst <- file.path(destDir, urlAndName[2])
        if(!file.exists(dst)){
          download.file(urlAndName[1], dst, mode="wb")
        }
      }
      
      apply(depSet, 2, dPairJar)
    }
  } else {
    sharedPaths <- ifFoundPaths
  }
  sharedPaths
}
