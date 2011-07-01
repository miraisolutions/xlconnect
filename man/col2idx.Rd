\name{col2idx}
\alias{col2idx}
\title{
Converting Excel column names to indices
}
\description{
Converts Excel column names to indices.
}
\usage{
col2idx(x)
}
\arguments{
  \item{x}{Character vector of Excel column names (e.g. "A", "AF", ...)
}
}
\value{
Returns a vector of integers representing the corresponding column indices.
Note that passing invalid column name references may result in an arbitrary number.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\link{idx2col}}, \code{\link{cref2idx}}, \code{\link{idx2cref}},
\code{\link{idx2aref}}, \code{\link{aref2idx}}, \code{\link{aref}}
}
\examples{
col2idx(c("A", "BTG"))
}
\keyword{utilities}
