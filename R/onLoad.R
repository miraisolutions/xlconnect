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
  poiPaths <- xlcEnsureDependenciesFor(cbind(
      c("https://repo1.maven.org/maven2/org/apache/poi/poi-ooxml/4.1.1/poi-ooxml-4.1.1.jar", "poi-ooxml.jar"),
      c("https://repo1.maven.org/maven2/org/apache/poi/poi/4.1.1/poi-4.1.1.jar", "poi.jar"),
      c("https://repo1.maven.org/maven2/org/apache/poi/poi-ooxml-schemas/4.1.1/poi-ooxml-schemas-4.1.1.jar", "poi-ooxml-schemas.jar")
    ), "poi-4\\.[1-9].*",
    c("/usr/share/java/poi.jar", "/usr/share/java/poi-ooxml.jar", "/usr/share/java/poi-ooxml-schemas.jar"), libname, pkgname)
  compressPaths <- xlcEnsureDependenciesFor(cbind(
      c("https://repo1.maven.org/maven2/org/apache/commons/commons-compress/1.19/commons-compress-1.19.jar", "commons-compress.jar")
    ), "commons-compress-1\\.(1[8-9]|[2-9][0-9]).*",
    c("/usr/share/java/commons-compress.jar"), libname, pkgname)
  xmlPath <- xlcEnsureDependenciesFor(cbind(
      c("https://repo1.maven.org/maven2/org/apache/xmlbeans/xmlbeans/3.1.0/xmlbeans-3.1.0.jar", "xmlbeans.jar")
  ), "xmlbeans-3\\..*",
    c("/usr/share/java/xmlbeans.jar"), libname, pkgname)
  collectionsPath <- xlcEnsureDependenciesFor(cbind(
      c("https://repo1.maven.org/maven2/org/apache/commons/commons-collections4/4.4/commons-collections4-4.4.jar", "commons-collections4.jar")
    ), "commons-collections4-4\\.([2-9]|1[0-9]).*",
    c("/usr/share/java/commons-collections4.jar"), libname, pkgname)
  mathPath <- xlcEnsureDependenciesFor(cbind(
      c("https://repo1.maven.org/maven2/org/apache/commons/commons-math3/3.6.1/commons-math3-3.6.1.jar", "commons-math3.jar")
    ), "commons-math3-3\\.([6-9]|1[0-9]).*",
    c("/usr/share/java/commons-math3.jar"), libname, pkgname)
  codecPath <- xlcEnsureDependenciesFor(cbind(
      c("https://repo1.maven.org/maven2/commons-codec/commons-codec/1.13/commons-codec-1.13.jar", "commons-math3.jar")
    ), "commons-codec-1\\.(1[1-9]|[2-9][0-9]).*",
    c("/usr/share/java/commons-codec.jar"), libname, pkgname)
  ooxmlSchemasPath <- xlcEnsureDependenciesFor(cbind(
    c("https://repo1.maven.org/maven2/org/apache/poi/ooxml-schemas/1.4/ooxml-schemas-1.4.jar", "ooxml-schemas.jar")
  ), "ooxml-schemas-1\\.([4-9]|[1-9][0-9]).*",
  c("/usr/share/java/ooxml-schemas.jar"), libname, pkgname)

  sharedPaths <- c(poiPaths, compressPaths, xmlPath, collectionsPath, mathPath, codecPath)
	.jpackage(name = pkgname, jars = "*", morePaths = sharedPaths)
	print(rJava::.jclassPath())
  
	# Perform general XLConnect settings - pass package description
	XLConnectSettings(packageDescription(pkgname))
}
