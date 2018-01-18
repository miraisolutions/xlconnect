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
# Get an existing cell style or create a new one if it does not exist yet.
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("getOrCreateCellStyle",
           function(object, name) standardGeneric("getOrCreateCellStyle"))

setMethod("getOrCreateCellStyle", 
          signature(object = "workbook", name = "character"), 
          function(object, name) {
            if(existsCellStyle(object, name)) {
              getCellStyle(object, name)
            } else {
              createCellStyle(object, name)
            }
          }
)
