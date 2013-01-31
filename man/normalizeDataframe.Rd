\name{normalizeDataframe}
\alias{normalizeDataframe}
\title{
Data frame Normalization for Unit Tests
}
\description{
Normalizes a data.frame with respect to column types. This is necessary for comparing data.frames written to Excel and then read from Excel. Excel does not know e.g. factor types. Factor variables
in fact are written to Excel as ordinary strings. Therefore, when read back in to R, they are treated as character variables.
}
\usage{
normalizeDataframe(df)
}
\arguments{
  \item{df}{\code{data.frame} to normalize}
}
\value{
Returns a normalized \code{data.frame}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\keyword{utilities}
\keyword{internal}
