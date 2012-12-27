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
# Saving Microsoft Excel workbooks
#
# A workbook's underlying Excel file is not saved (or being created in case the 
# file did not exist and create = TRUE has been specified) unless the saveWorkbook
# method has been called on the object. This provides more flexibility to the user 
# to decide when changes are saved and also provides better performance in that 
# several changes can be written in one go (normally at the end, rather than after 
# every operation causing the file to be rewritten again completely each time). This 
# is due to the fact that workbooks are manipulated in-memory and are only written 
# to disk with specifically calling saveWorkbook.
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

setGeneric("saveWorkbook",
	function(object, file) standardGeneric("saveWorkbook"))

setMethod("saveWorkbook", signature(object = "workbook", "missing"), function(object, file) {
	jTryCatch(object@jobj$save())
})

setMethod("saveWorkbook", signature(object = "workbook", "character"), function(object, file) {
	jTryCatch(object@jobj$save(path.expand(file)))
})
