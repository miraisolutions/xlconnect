#############################################################################
#
# XLConnect
# Copyright (C) 2010-2013 Mirai Solutions GmbH
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
# Setting the cell style for the DATATYPE style action
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("setCellStyleForType",
           function(object, type, style) standardGeneric("setCellStyleForType"))

setMethod("setCellStyleForType", 
          signature(object = "workbook"), 
          function(object, type, style) {
            xlcCall(object, "setCellStyleForDataType", classToXlcType(type), style)
            invisible()
          }
)