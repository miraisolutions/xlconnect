\name{xlcFreeMemory}
\alias{xlcFreeMemory}
\title{
Freeing Java Virtual Machine memory
}
\description{
Frees Java Virtual Machine (JVM) memory.
}
\usage{
xlcFreeMemory(...)
}
\arguments{
  \item{\dots}{
  Further arguments to be passed to R's garbage collector (\code{\link{gc}}).
}
}
\details{
This function uses Java's Runtime class to run the garbage collector.\cr
Java memory is freed by first running R's garbage collector (\code{\link{gc}}) and then
Java's garbage collector. This sequence is important as R's \code{\link{gc}} may release 
objects which in turn allows Java's garbage collector to release some objects.\cr\cr
Note, in general there should be no need to make active use of this with \pkg{XLConnect}. Both
R and Java automatically perform garbage collection at times. However, this function might be useful
to reclaim memory after removing a large data object that has been written/read to/from Excel.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{xlcMemoryReport}}, \code{\link{gc}}
}
\examples{
\dontrun{
xlcFreeMemory()
}
}
\keyword{utilities}
