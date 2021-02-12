\name{idx2cref}
\alias{idx2cref}
\title{
Converting indices to Excel cell references
}
\description{
Converts row & column indices to Excel cell references
}
\usage{
idx2cref(x, absRow = TRUE, absCol = TRUE)
}
\arguments{
  \item{x}{Numeric (integer) matrix or vector of indices. If a matrix is provided it
  should have two columns with the first column representing the row indices and the second
  column representing the column indices (i.e. each row represents a index-based cell reference). 
  If a vector is provided it will be converted to a matrix by filling the vector into a 2-column 
  matrix by row.
}
  \item{absRow}{Boolean determining if the row index should be considered absolute. 
  If \code{TRUE} (default), this will result in a '$'-prefixed row identifier.
}
  \item{absCol}{Boolean determining if the column index should be considered absolute.
  If \code{TRUE} (default), this will result in a '$'-prefixed column identifier.
}
}
\value{
Returns a character vector of corresponding Excel cell references.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{cref2idx}}, \code{\link{idx2col}}, \code{\link{col2idx}},
\code{\link{idx2aref}}, \code{\link{aref2idx}}, \code{\link{aref}}
}
\examples{
\dontrun{
idx2cref(c(5, 8, 14, 38))
}
}
\keyword{utilities}
