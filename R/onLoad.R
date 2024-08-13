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
# XLConnect Package Initialization
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

.onLoad <- function(libname, pkgname) {
  
  javaCheck <- function() {
    # Java version check, without initializing rJava (we can only do .jpackage(...) after downloading resources below)
    rawVersion <- system2("java", c("-version"), stdout = TRUE, stderr = TRUE)
    versionLines <- rawVersion[!grepl("Picked up", rawVersion)] # ignore java environment variables being logged
    jv <- regmatches(versionLines[1], regexpr("[0-9\\.]+", versionLines[1]))
    if (nchar(jv)<1) {
      stop(paste0("unable to parse java version from", paste0(versionLines, collapse=" "), "; is java installed correctly ?"))
    }
    else if(nchar(jv)==1) {
       jvn <- as.numeric(jv)
    }
    else {
      twoFirst <- substr(jv, 1L, 2L)
      if(twoFirst == "1.") {
        jvn <- as.numeric(substr(jv,3L,3L))
      } else {
        jvn <- as.numeric(twoFirst)
      }
    }
    if (jvn<8) stop(paste0("XLConnect is compatible with Java versions 8 and above. Detected java version: ",jv))
  }
  javaCheck()
  repo <- Sys.getenv("XLCONNECT_JAVA_REPO_URL")
  if (is.null(repo) || repo=="") {
    repo <- "https://repo1.maven.org/maven2"
  }
  apachePrefix <- paste0(repo, "/org/apache")
  sharedPaths <- tryCatch({
    c(
      xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/poi-ooxml-full/5.3.0/poi-ooxml-full-5.3.0.jar"), "poi-ooxml-full.jar", 
      "5.3.0",  libname, pkgname),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/poi-ooxml/5.3.0/poi-ooxml-5.3.0.jar"), "poi-ooxml.jar", 
      "5.3.0",  libname, pkgname, debianpkg = "libapache-poi-java", rpmpkg="apache-poi"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/poi/poi/5.3.0/poi-5.3.0.jar"), "poi.jar", 
      "5.3.0",  libname, pkgname, debianpkg = "libapache-poi-java", rpmpkg="apache-poi"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-compress/1.26.2/commons-compress-1.26.2.jar"), "commons-compress.jar",
      "1\\.(2[5-9]|[2-9][0-9]).*",  libname, pkgname, debianpkg = "libcommons-compress-java", rpmpkg="apache-commons-compress"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-lang3/3.14.0/commons-lang3-3.14.0.jar"), "commons-lang3-3.14.0.jar",
      "3\\.(1[4-9]|[2-9][0-9])\\.*",  libname, pkgname, debianpkg="libcommons-lang3-java", rpmpkg="apache-commons-lang3"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/xmlbeans/xmlbeans/5.2.1/xmlbeans-5.2.1.jar"), "xmlbeans.jar",
      "5\\..*",  libname, pkgname, debianpkg="libxmlbeans-java"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-collections4/4.4/commons-collections4-4.4.jar"), "commons-collections4.jar",
      "4-4\\.([2-9]|1[0-9]).*",  libname, pkgname, debianpkg="libcommons-collections4-java", rpmpkg="apache-commons-collections4"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/commons/commons-math3/3.6.1/commons-math3-3.6.1.jar"), "commons-math3.jar",
      "3\\.([6-9]|1[0-9]).*",  libname, pkgname, debianpkg="libcommons-math3-java"),
    xlcEnsureDependenciesFor(
      paste0(apachePrefix, "/logging/log4j/log4j-api/2.23.1/log4j-api-2.23.1.jar"), "log4j-api.jar",
      "2\\.23\\..*",  libname, pkgname),
    xlcEnsureDependenciesFor(
      paste0(repo, "/commons-codec/commons-codec/1.17.0/commons-codec-1.17.0.jar"), "commons-codec-1.17.0.jar",
      "1\\.(1[1-9]|[2-9][0-9]).*",  libname, pkgname, debianpkg="libcommons-codec-java", rpmpkg="apache-commons-codec"),
    xlcEnsureDependenciesFor(
      paste0(repo, "/commons-io/commons-io/2.16.1/commons-io-2.16.1.jar"), "commons-io-2.16.1.jar",
      "2\\.(1[5-9]|[2-9][0-9]).*",  libname, pkgname, debianpkg="libcommons-io-java", rpmpkg="apache-commons-io"),
    xlcEnsureDependenciesFor(
      paste0(repo, "/com/zaxxer/SparseBitSet/1.3/SparseBitSet-1.3.jar"), "SparseBitSet.jar",
      "1\\.([2-9]|[1-9][0-9]).*",  libname, pkgname)
    )
  },
  error=function(e) {
          e
        }
  )
  .jpackage(name = pkgname, jars = "*", morePaths = sharedPaths, own.loader=TRUE)  
  # Perform general XLConnect settings - pass package description
  XLConnectSettings(packageDescription(pkgname))
}
