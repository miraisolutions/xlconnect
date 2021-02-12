\name{idx2col}
\alias{idx2col}
\title{
Converting column indices to Excel column names
}
\description{
Converts column indices to Excel column names.
}
\usage{
idx2col(x)
}
\arguments{
  \item{x}{Numeric (integer) vector of column indices
}
}
\value{
Returns a character vector of corresponding Excel column names.
Numbers \code{<= 0} result in the empty string ("").
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{col2idx}}, \code{\link{idx2cref}}, \code{\link{cref2idx}},
\code{\link{idx2aref}}, \code{\link{aref2idx}}, \code{\link{aref}}
}
\examples{
\dontrun{
idx2col(c(1, 347))
}
}
\keyword{utilities}
