#############################################################################
#
# XLConnect
# Copyright (C) 2010-2021 Mirai Solutions GmbH
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
# XLConnect Package Initialization
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

.onLoad <- function(libname, pkgname) {
  
  javaCheck <- function() {
    # Java version check, without .jinit (we do .jpackage after downloading resources)
    rawVersion <- system2("java", c("-version"), stdout = TRUE, stderr = TRUE)
    jv <- regmatches(rawVersion[1], regexpr("[0-9]+\\.[0-9\\.]*", rawVersion[1]))
    twoFirst <- substr(jv, 1L, 2L)
    if(twoFirst == "1.") {
      jvn <- as.numeric(substr(jv,3L,3L))
    } else {
      jvn <- as.numeric(twoFirst)
    }
    if (jvn<8 || jvn>15) stop(paste0("XLConnect is compatible with Java versions 8 to 15. Detected java version: ",jv))
  }
  javaCheck()
  repo <- Sys.getenv("XLCONNECT_JAVA_REPO_URL")
  if (is.null(repo) || repo=="") {
    repo <- "https://repo1.maven.org/maven2"
  }
  apachePrefix <- paste0(repo, "/org/apache")
  sharedPaths <- tryCatch({
    c(xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/poi-ooxml-schemas/4.1.2/poi-ooxml-schemas-4.1.2.jar"), "poi-ooxml-schemas.jar", 
      "4\\.[1-9].*",  libname, pkgname, debianpkg = "libapache-poi-java", rpmpkg="apache-poi"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/poi-ooxml/4.1.2/poi-ooxml-4.1.2.jar"), "poi-ooxml.jar", 
      "4\\.[1-9].*",  libname, pkgname, debianpkg = "libapache-poi-java", rpmpkg="apache-poi"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/poi/4.1.2/poi-4.1.2.jar"), "poi.jar", 
      "4\\.[1-9].*",  libname, pkgname, debianpkg = "libapache-poi-java", rpmpkg="apache-poi"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-compress/1.20/commons-compress-1.20.jar"), "commons-compress.jar",
      "1\\.(1[8-9]|[2-9][0-9]).*",  libname, pkgname, debianpkg = "libcommons-compress-java", rpmpkg="apache-commons-compress"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/xmlbeans/xmlbeans/3.1.0/xmlbeans-3.1.0.jar"), "xmlbeans.jar",
      "3\\..*",  libname, pkgname, debianpkg="libxmlbeans-java"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-collections4/4.4/commons-collections4-4.4.jar"), "commons-collections4.jar",
      "4-4\\.([2-9]|1[0-9]).*",  libname, pkgname, debianpkg="libcommons-collections4-java", rpmpkg="apache-commons-collections4"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-math3/3.6.1/commons-math3-3.6.1.jar"), "commons-math3.jar",
      "3\\.([6-9]|1[0-9]).*",  libname, pkgname, debianpkg="libcommons-math3-java"),
    xlcEnsureDependenciesFor(
      paste0(repo, "/commons-codec/commons-codec/1.15/commons-codec-1.15.jar"), "commons-codec-1.15.jar",
      "1\\.(1[1-9]|[2-9][0-9]).*",  libname, pkgname, debianpkg="libcommons-codec-java", rpmpkg="apache-commons-codec"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/ooxml-schemas/1.4/ooxml-schemas-1.4.jar"), "ooxml-schemas.jar",
      "1\\.([4-9]|[1-9][0-9]).*",  libname, pkgname))
  },
  error=function(e) {
          write("downloading JAR dependencies failed!", file.path(libname, pkgname, ".fail"))
          e
        }
  )
  .jpackage(name = pkgname, jars = "*", morePaths = sharedPaths)
  # Perform general XLConnect settings - pass package description
  XLConnectSettings(packageDescription(pkgname))
}
