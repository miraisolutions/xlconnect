\name{aref}
\alias{aref}
\title{
Constructing Excel area references
}
\description{
Constructs an Excel area reference
}
\usage{
aref(topLeft, dimension)
}
\arguments{
  \item{topLeft}{Top left corner. Either a \code{character} specifying a cell reference in the
  form "A1" or a \code{numeric} vector of length two specifying the corresponding coordinates.}
  \item{dimension}{Dimensions (\code{numeric}) of a 2-dimensional object 
  (mostly a \code{data.frame} or a \code{matrix})}
}
\value{
Returns the area reference (\code{character}) for the specified top left cell and dimension.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\link{aref2idx}}, \code{\link{idx2aref}},
\code{\link{idx2cref}}, \code{\link{col2idx}}, \code{\link{idx2col}}
}
\examples{
aref("A1", dim(mtcars))
aref(c(1, 1), dim(mtcars))
}
\keyword{utilities}
