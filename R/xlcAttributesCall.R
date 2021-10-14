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

xlcAttributesCall <-
  function(obj,
           fun,
           ...,
           .recycle = TRUE,
           .simplify = TRUE) {
    res <- xlcCall(obj, fun, ..., .recycle = TRUE, .simplify = TRUE)
    definedNames <- names(res)[!is.na(names(res))]
    names(res) <- rep(definedNames, length.out = length(names(res)))
    
    jni <- mapply(function(resw) {
      .jcall(resw, "S", "jni")
    }, res)
    
    unwrapped <- mapply(function(resw, jtype) {
      .jcall(resw, jtype, "getValue")
    }, res, jni)
    
    allANames = unique(Reduce(function(resw1, resw2) {
      append(
        .jcall(resw1, "[S", "getAttributeNames"),
        .jcall(resw2, "[S", "getAttributeNames")
      )
    }, res))
    
    
    attributeRows <- Map(function(aName) {
      Map(function(resw) {
        thev <- .jcall(resw, "S", "getAttributeValue", aName)
        if (is.null(thev))
          NA
        else
          thev
      }, res)
    }, allANames)
    
    attrMtx <- Reduce(rbind, attributeRows)
    if (length(allANames) > 1) {
      colnames(attrMtx) <- allANames
      
      Map(function(aName) {
        attr(unwrapped, aName) <- attrMtx[, aName]
      }, allANames)
    }
    else
      attr(unwrapped, allANames[1]) <- attrMtx
    unwrapped
  }