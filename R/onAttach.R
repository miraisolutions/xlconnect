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

.onAttach <- function(libname, pkgname) {
  # Print package information
  pdesc <- packageDescription(pkgname)
  if(file.exists(file.path(libname, pkgname, ".fail"))){
    packageStartupMessage(paste0("XLConnect: It seems downloading the JAR dependencies may have failed. ",
                   "If you would like to use a different maven repository, ",
                   "please set the environment variable XLCONNECT_JAVA_REPO_URL to a valid URL, ",
                   "e.g. Sys.setenv(XLCONNECT_JAVA_REPO_URL='https://jcenter.bintray.com'), ",
                   "and reinstall the package."))
  }
  packageStartupMessage(pdesc$Package, " ", pdesc$Version, " by ", pdesc$Author)
  packageStartupMessage(pdesc$URL)
}
