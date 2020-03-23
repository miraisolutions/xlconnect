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
# XLConnect Package Initialization
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

.onLoad <- function(libname, pkgname) {
  repo <- Sys.getenv("XLCONNECT_JAVA_REPO_URL")
  if (is.null(repo) || repo=="") {
    repo <- "https://repo1.maven.org/maven2"
  }
  apachePrefix <- paste0(repo, "/org/apache")
  sharedPaths <- tryCatch({
    c(xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/poi-ooxml/4.1.1/poi-ooxml-4.1.1.jar"), "poi-ooxml.jar", 
      "poi-ooxml-4\\.[1-9].*",  libname, pkgname),
     xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/poi/4.1.1/poi-4.1.1.jar"), "poi.jar", 
      "poi-4\\.[1-9].*",  libname, pkgname),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-compress/1.19/commons-compress-1.19.jar"), "commons-compress.jar",
      "commons-compress-1\\.(1[8-9]|[2-9][0-9]).*",  libname, pkgname),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/xmlbeans/xmlbeans/3.1.0/xmlbeans-3.1.0.jar"), "xmlbeans.jar",
      "xmlbeans-3\\..*",  libname, pkgname),
    xlcEnsureDependenciesFor(
        paste0(apachePrefix, "/commons/commons-collections4/4.4/commons-collections4-4.4.jar"), "commons-collections4.jar",
        "commons-collections4-4\\.([2-9]|1[0-9]).*",  libname, pkgname),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-math3/3.6.1/commons-math3-3.6.1.jar"), "commons-math3.jar",
      "commons-math3-3\\.([6-9]|1[0-9]).*",  libname, pkgname),
    xlcEnsureDependenciesFor(
      paste0(repo, "/commons-codec/commons-codec/1.13/commons-codec-1.13.jar"), "commons-math3.jar",
      "commons-codec-1\\.(1[1-9]|[2-9][0-9]).*",  libname, pkgname),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/ooxml-schemas/1.4/ooxml-schemas-1.4.jar"), "ooxml-schemas.jar",
      "ooxml-schemas-1\\.([4-9]|[1-9][0-9]).*",  libname, pkgname))
  },
  error=function(e) {
          write("downloading JAR dependencies failed!", file.path(libname, pkgname, ".fail"))
          e
        }
  )
  .jpackage(name = pkgname, jars = "*", morePaths = sharedPaths)
  # Java version check
  jv <- .jcall("java/lang/System", "S", "getProperty", "java.runtime.version")
  twoFirst <- substr(jv, 1L, 2L)
  if(twoFirst == "1.") {
  	jvn <- as.numeric(substr(jv,3L,3L))
  } else {
	  jvn <- as.numeric(twoFirst)
  }
  if (jvn<8 || jvn>11) stop(paste0("Installed java version ",jv ," is not between Java>=8 and <=11! This is needed for this package"))

  # Perform general XLConnect settings - pass package description
  XLConnectSettings(packageDescription(pkgname))
}
