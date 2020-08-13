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
# XLConnect Package Installation: find system packages by name, including
# installed files, using debian and / or rpm package managers
# 
# example: with 
#
# Author: Simon Poltier, Mirai Solutions GmbH
#
#############################################################################

 checkSystemPackage <- function (debianpkgname, rpmpkgname, versionpattern) {
  dpkg <- function(args) {
    system2("dpkg", args, stdout = TRUE)
  }
  rpm <- function(args) {
    system2("rpm", args, stdout = TRUE)
  }
  if(!is.null(debianpkgname) && system2("dpkg", c("--help"), stdout=FALSE) == 0) {
    pkgList <- dpkg(c("-l", debianpkgname))
    pkgLine <- pkgList[which(grepl(debianpkgname, pkgList))][1]
    foundVersion <- strsplit(pkgLine, " +")[[1]][3]
    if(grepl(versionpattern, foundVersion)) {
      allFiles <- dpkg(c("--listfiles", debianpkgname))
      allFiles[which(grepl(".*/java.*jar", allFiles))]
    } else { c() }
  }
  else if (!is.null(rpmpkgname) && system2("rpm", c("--help"), stdout = FALSE) == 0) {
    pkgInfo <- rpm(c("-qi", rpmpkgname))
    versionStr <- pkgInfo[which(grepl('Version', pkgInfo))]
    foundVersion <- strsplit(versionStr, " +: +")[[1]][2]
    if(grepl(versionpattern, foundVersion)) {
      allFiles <- rpm(c("-ql", rpmpkgname))
      allFiles[which(grepl(".*/java.*jar", allFiles))]
    } else { c() }
  }
  else { c() }
}