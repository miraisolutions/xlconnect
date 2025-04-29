#############################################################################
#
# XLConnect
# Copyright (C) 2010-2025 Mirai Solutions GmbH
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
# Utility function for converting XLConnect java objects with added attributes
# to R variables with the corresponding attributes.
# Sets the attributes if option XLConnect.setCustomAttributes is TRUE. 
# Otherwise only unwraps the java object.
#
# Author: Simon Poltier, Mirai Solutions GmbH
#
#############################################################################

withAttributesFromJava <- function(jobj) {
    
    unwrapped <- jobj$getValue()
    allANames = .jcall(jobj, "[S", "getAttributeNames")
    
    if(getOption("XLConnect.setCustomAttributes")){
        for(i in seq(along = allANames)) {
            attr(unwrapped, allANames[i]) <- jobj$getAttributeValue(allANames[i])
        }
    }
    unwrapped
  }