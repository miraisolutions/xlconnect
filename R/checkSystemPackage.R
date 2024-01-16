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
# XLConnect Package Installation: find system packages by name, including
# installed files, using Debian or Red Hat package managers
#
# Author: Simon Poltier, Mirai Solutions GmbH
#
#############################################################################

 checkSystemPackage <- function (debianpkgname, rpmpkgname, versionpattern) {
  dpkg <- function(args) {
    suppressWarnings(system2("dpkg", args, stdout = TRUE, stderr = TRUE))
  }
  rpm <- function(args) {
    suppressWarnings(system2("rpm", args, stdout = TRUE, stderr = TRUE))
  }
  distro <- function(distroName) {
    releaseLines <- suppressWarnings(system2("cat",c("/etc/*-release"), stdout = TRUE, stderr = TRUE))
    sum(grepl(pattern = paste0("ID.*",distroName), x = releaseLines, ignore.case = TRUE), na.rm = TRUE) > 0
  }
  if (!is.null(rpmpkgname) && distro("rhel")) {
    pkgInfo <- rpm(c("-qi", rpmpkgname))
    versionStr <- pkgInfo[which(grepl('Version', pkgInfo))]
    if(length(versionStr) != 0) {
      foundVersion <- strsplit(versionStr, " +: +")[[1]][2]
      if(grepl(versionpattern, foundVersion)) {
        allFiles <- rpm(c("-ql", rpmpkgname))
        allFiles[which(grepl(".*/java.*jar", allFiles))]
      } else { c() }
    } else { c() }
  }
  else if(!is.null(debianpkgname) && distro("debian")) {
    pkgList <- dpkg(c("-l", debianpkgname))
    pkgLine <- pkgList[which(grepl(debianpkgname, pkgList))][1]
    foundVersion <- strsplit(pkgLine, " +")[[1]][3]
    if(grepl(versionpattern, foundVersion)) {
      allFiles <- dpkg(c("--listfiles", debianpkgname))
      allFiles[which(grepl(".*/java.*jar", allFiles))]
    } else { c() }
  }
  else { c() }
 }
 