\name{configurePOI}
\alias{configurePOI}
\title{
Configuring Apache POI
}
\description{
Configures Apache POI and related components.
}
\usage{
configurePOI(zip_max_files = 1000L, zip_min_inflate_ratio = 0.001,
  zip_max_entry_size = 0xFFFFFFFF, zip_max_text_size = 10*1024*1024,
  zip_entry_threshold_bytes = -1L, max_size_byte_array = -1L,
  java_io_tmpdir = tempdir())
}
\arguments{
  \item{zip_max_files}{Integer scalar specifying the maximum number of files
  allowed inside an *.xlsx file. Defaults to \code{1000}.}
  \item{zip_min_inflate_ratio}{Numeric scalar specifying the ratio between
  de- and inflated bytes to detect zip-bombs. If the compression ratio is
  better than the specified number an error will be thrown. Defaults to
  \code{0.001}.}
  \item{zip_max_entry_size}{Integer scalar specifying the maximum file size
  of a single zip entry in an *.xlsx file. Defaults to 4'294'967'295 bytes,
  which is 4GB.}
  \item{zip_max_text_size}{Integer scalar specifying the maximum number of
  characters of text that are extracted before an error is thrown. Defaults
  to 10'485'760.}
  \item{zip_entry_threshold_bytes}{Integer scalar specifying the number of bytes
  at which a zip entry is regarded as too large for holding in memory and the
  data is put in a temp file instead. Defaults to \code{-1L}, meaning temp files
  are not used and that zip entries with more than 2GB of data after decompressing
  will fail. \code{0L} means all zip entries are stored in temp files.}
  \item{max_size_byte_array}{Integer scalar specifying the maximum number of
  bytes that should be possible to be allocated in a single step. Increasing
  this limit can help if you are dealing with large Excel files, but note that
  this may demand a larger heap space (see option \code{java.parameters}; e.g.
  \code{options(java.parameters = "-Xmx8192m")}. Defaults to -1, which means
  that record-specific limits apply.)}
  \item{java_io_tmpdir}{Directory to hold POI temporary files and generally any
  temporary files produced by the Java subprocess. Defaults to the R session
  temporary directory \code{tempdir()}}
}
\details{
Many of the settings exposed here exist for security reasons to prevent excessive
memory consumption and protect against security vulnerabilities when processing
documents provided by untrusted sources.
}
\references{
Apache POI configuration: \url{https://poi.apache.org/components/configuration.html} 
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\examples{
\dontrun{
configurePOI(zip_max_files = 5000L, max_size_byte_array = 250000000L)
}
}
\keyword{IO}
