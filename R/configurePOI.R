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
# Configures Apache POI and related components.
# See https://poi.apache.org/components/configuration.html
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

configurePOI <- function(
    zip_max_files = 1000L,
    zip_min_inflate_ratio = 0.001,
    zip_max_entry_size = 0xFFFFFFFF,
    zip_max_text_size = 10*1024*1024,
    zip_entry_threshold_bytes = -1L,
    max_size_byte_array = -1L,
    java_io_tmpdir = tempdir()
) {
  J("java.lang.System")$setProperty("java.io.tmpdir", java_io_tmpdir)
  
  ioutils <- J("org.apache.poi.util.IOUtils")
  ioutils$setByteArrayMaxOverride(as.integer(max_size_byte_array))
  
  zip <- J("org.apache.poi.openxml4j.util.ZipSecureFile")
  zip$setMaxFileCount(rJava::.jlong(zip_max_files))
  zip$setMinInflateRatio(zip_min_inflate_ratio)
  zip$setMaxEntrySize(rJava::.jlong(zip_max_entry_size))
  zip$setMaxTextSize(rJava::.jlong(zip_max_text_size))
  
  zipEntry <- J("org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource")
  zipEntry$setThresholdBytesForTempFiles(as.integer(zip_entry_threshold_bytes))
  
  invisible()
}
