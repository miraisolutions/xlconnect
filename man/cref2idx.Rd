\name{cref2idx}
\alias{cref2idx}
\title{
Converting Excel cell references to indices
}
\description{
Converts Excel cell references to row & column indices
}
\usage{
cref2idx(x)
}
\arguments{
  \item{x}{Character vector of Excel cell references (e.g. "$A$20", "B18", ...)
}
}
\value{
Returns a numeric matrix with two columns and as many rows as cell references that have
been provided. The first column represents the row indices and the second column represents
the column indices.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\link{idx2cref}}, \code{\link{col2idx}}, \code{\link{idx2col}},
\code{\link{idx2aref}}, \code{\link{aref2idx}}, \code{\link{aref}}
}
\examples{
cref2idx(c("$A$20", "B18"))
}
\keyword{utilities}
