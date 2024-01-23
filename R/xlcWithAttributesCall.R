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
# Extracts attributes from a wrapper Java type and sets them on the R object.
# 
# Author: Martin Studer & Simon Poltier, Mirai Solutions GmbH
#
#############################################################################

xlcWithAttributesCall <- function(obj, fun, ..., .recycle = TRUE, .simplify = TRUE, .checkWarnings = TRUE) {
  g = eval(parse(text = paste("obj@jobj$", fun, sep = "")))
  f <- function(...) withAttributesFromJava(g(...))
  args <- list(...)
  if(.recycle) {
    args <- lapply(args, function(x) {
      if(is.atomic(x)) x
      else wrapList(x)
    })
    res = jTryCatch(do.call("mapply", args = c(FUN = f, args, SIMPLIFY = FALSE, USE.NAMES = FALSE)))
    if(.simplify) {
      res_attr <- Reduce(function(atts1, atts2) {
        aNames <- unique(c(names(atts1), names(atts2)))
        sapply(aNames, function(aName) { list(c(atts1[aName][[1]], atts2[aName][[1]])) })
      } ,lapply(res, attributes))
      res <- simplify2array(res)
      attributes(res) <- res_attr
    }
  } else {
    res = jTryCatch(do.call(f, args))
  }

  if (.checkWarnings) {
	  warnings = .jcall(obj@jobj, "[S", "retrieveWarnings")
	  for(w in warnings) warning(w, call. = FALSE)
	}
  
  res
}
