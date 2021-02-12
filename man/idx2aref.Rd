\name{idx2aref}
\alias{idx2aref}
\title{
Converting row and column based area references to Excel area references
}
\description{
Converts row & column based area references to Excel area references
}
\usage{
idx2aref(x)
}
\arguments{
  \item{x}{Numeric (integer) matrix or vector of indices. If a matrix is provided it
  should have four columns with the first two columns representing the top left corner
  (row and column indices) and the third & fourth column representing the bottom right
  corner. If a vector is provided it will be converted to a matrix by filling the vector into a 4-column 
  matrix by row.}
}
\value{
Returns a \code{character} vector of corresponding Excel area references.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{aref2idx}}, \code{\link{aref}}, \code{\link{idx2cref}},
\code{\link{cref2idx}}, \code{\link{idx2col}}, \code{\link{col2idx}}
}
\examples{
\dontrun{
idx2aref(c(1, 1, 5, 4))
}
}
\keyword{utilities}
