\name{aref2idx}
\alias{aref2idx}
\title{
Converting Excel cell references to row and column based cell references
}
\description{
Converts Excel cell references to row and column based cell references
}
\usage{
aref2idx(x)
}
\arguments{
  \item{x}{\code{Character} vector of Excel cell references (e.g. "A1:B6", "B6:C17", ...)
}
}
\value{
Returns a numeric matrix with four columns and as many rows as cell references that have
been provided. The first two columns represent the coordinates of the top left corner (row, column)
and the third and fourth columns represent the bottom right corner of the referenced area.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{idx2aref}}, \code{\link{aref}}, \code{\link{cref2idx}},
\code{\link{idx2cref}}, \code{\link{col2idx}}, \code{\link{idx2col}}
}
\examples{
\dontrun{
aref2idx(c("A1:B6", "B6:C17"))
}
}
\keyword{utilities}
