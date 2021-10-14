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
# Utility function for vectorizing argument lists and default Java exception
# handling (jTryCatch)
#
# Atomic objects are replicated as is, others are wrapped in a list as defined
# by wrapList and then replicated
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

xlcAttributesCall <- function(obj, fun, ..., .recycle = TRUE, .simplify = TRUE) {
  res <- xlcCall(obj, fun, ..., .recycle = TRUE, .simplify = TRUE)
  definedNames <- names(res)[!is.na(names(res))]
  names(res) <- rep(definedNames, length.out = length(names(res)))
  jni <- mapply(function(resw) {
    .jcall(resw, "S", "jni")
  }, res)
  unwrapped <- mapply(function(resw, jtype) {
    .jcall(resw, jtype, "getValue")
  }, res, jni)
  
  attributesToSet <- mapply(function(resw) {
    aNames <- .jcall(resw, "[S", "getAttributeNames")
    aValues <- .jcall(resw, "[S", "getAttributeValues")
    attList <- list()
    for (i in seq(along = aNames)) {
      attList[aNames[i]]=aValues[i]
    }
    attList
  }, res)
  # print(attributesToSet)
  attributes(unwrapped) <- append(attributes(unwrapped), attributesToSet)
  unwrapped
}