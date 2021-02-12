\name{xlcMemoryReport}
\alias{xlcMemoryReport}
\title{
Reporting free Java Virtual Machine memory
}
\description{
Reports the amount of free memory in the Java Virtual Machine (JVM).
}
\usage{
xlcMemoryReport()
}
\details{
This function uses Java's Runtime class to query the free JVM memory.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{xlcFreeMemory}}
}
\examples{
\dontrun{
xlcMemoryReport()
}
}
\keyword{utilities}
