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
# Specifying the font color for cell styles
#
# Author: Matt Deppe
#
#############################################################################

setGeneric("setFontColor",
           function(object, color) standardGeneric("setFontColor"))

setMethod("setFontColor",
          signature(object = "cellstyle", color = "ANY"),
          function(object, color) {
            if (is.numeric(color)) {
              # Pass as integer color to Java
              xlcCall(object, "setFontColor", as.integer(color), NULL,
                      .recycle = FALSE, .checkWarnings = FALSE)
            } else if (is.character(color) && grepl("^#([A-Fa-f0-9]{6})$", color)) {
              # Convert hex color to RGB and pass as raw to Java
              rgb <- as.raw(col2rgb(color))
              xlcCall(object, "setFontColor", as.integer(0), rgb,
                      .recycle = FALSE, .checkWarnings = FALSE)
            } else {
              stop("Invalid color format. Provide a numeric value or a valid hex color string (e.g., '#RRGGBB').")
            }
            invisible()
          }
)



